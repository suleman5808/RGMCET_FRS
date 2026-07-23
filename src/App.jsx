import { useState } from 'react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import './style.css';

const MASTER_ID_HEADINGS = [
  'staff id',
  'staffid',
  'faculty id',
  'facultyid',
  'employee id',
  'employeeid',
  'id',
];

const MASTER_NAME_HEADINGS = [
  'faculty name',
  'facultyname',
  'staff name',
  'staffname',
  'name',
];

const MASTER_DEPT_HEADINGS = [
  'department',
  'dept',
  'branch',
  'faculty department',
  'staff department',
];

const ATTENDANCE_ID_HEADINGS = [
  'staff id',
  'staffid',
  'faculty id',
  'facultyid',
  'employee id',
  'employeeid',
  'id',
];

const ATTENDANCE_HEADINGS = ['attendance', 'attendance status', 'status'];

function normalizeHeading(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/[_-]/g, ' ')
    .replace(/\s+/g, ' ');
}

function normalizeValue(value) {
  if (value === undefined || value === null) {
    return '';
  }

  return String(value).trim().toUpperCase();
}

function cleanText(value) {
  if (value === undefined || value === null) {
    return '';
  }

  return String(value).trim();
}

function findColumnName(row, acceptedHeadings) {
  return Object.keys(row).find((column) =>
    acceptedHeadings.includes(normalizeHeading(column))
  );
}

async function readExcelFile(file) {
  const arrayBuffer = await file.arrayBuffer();

  const workbook = XLSX.read(arrayBuffer, {
    type: 'array',
    cellDates: false,
  });

  const allRows = [];

  workbook.SheetNames.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];

    const rows = XLSX.utils.sheet_to_json(worksheet, {
      defval: '',
      raw: false,
    });

    allRows.push(...rows);
  });

  return allRows;
}

function getReportDate() {
  return new Intl.DateTimeFormat('en-GB', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
  }).format(new Date());
}

function App() {
  const [masterFaculty, setMasterFaculty] = useState([]);
  const [masterFileName, setMasterFileName] = useState('');
  const [attendanceFileNames, setAttendanceFileNames] = useState([]);
  const [absentFaculty, setAbsentFaculty] = useState([]);
  const [unmatchedIds, setUnmatchedIds] = useState([]);
  const [message, setMessage] = useState(
    'Upload the master faculty Excel file first.'
  );
  const [loading, setLoading] = useState(false);

  const handleMasterFile = async (event) => {
    const file = event.target.files?.[0];

    if (!file) {
      return;
    }

    setLoading(true);
    setAbsentFaculty([]);
    setUnmatchedIds([]);
    setAttendanceFileNames([]);

    try {
      const rows = await readExcelFile(file);

      if (rows.length === 0) {
        throw new Error('The master Excel file does not contain faculty data.');
      }

      const firstValidRow = rows.find((row) => Object.keys(row).length > 0);

      if (!firstValidRow) {
        throw new Error('No valid data was found in the master file.');
      }

      const idColumn = findColumnName(firstValidRow, MASTER_ID_HEADINGS);

      const nameColumn = findColumnName(firstValidRow, MASTER_NAME_HEADINGS);

      const deptColumn = findColumnName(firstValidRow, MASTER_DEPT_HEADINGS);

      if (!idColumn) {
        throw new Error(
          'Staff ID column was not found. Use Staff ID or Faculty ID.'
        );
      }

      if (!nameColumn) {
        throw new Error(
          'Faculty Name column was not found. Use Faculty Name or Staff Name.'
        );
      }

      const faculty = rows
        .map((row) => ({
          staffId: normalizeValue(row[idColumn]),
          name: cleanText(row[nameColumn]),
          department: deptColumn ? cleanText(row[deptColumn]) : '',
        }))
        .filter((member) => member.staffId && member.name);

      if (faculty.length === 0) {
        throw new Error(
          'No valid Staff ID and Faculty Name records were found.'
        );
      }

      const uniqueFacultyMap = new Map();

      faculty.forEach((member) => {
        if (!uniqueFacultyMap.has(member.staffId)) {
          uniqueFacultyMap.set(member.staffId, member);
        }
      });

      const uniqueFaculty = Array.from(uniqueFacultyMap.values());

      setMasterFaculty(uniqueFaculty);
      setMasterFileName(file.name);

      setMessage(
        `${uniqueFaculty.length} faculty records loaded. Now upload the attendance file.`
      );
    } catch (error) {
      setMasterFaculty([]);
      setMasterFileName('');
      setMessage(error.message);
    } finally {
      setLoading(false);
      event.target.value = '';
    }
  };

  const handleAttendanceFiles = async (event) => {
    const selectedFiles = Array.from(event.target.files || []);

    if (selectedFiles.length === 0) {
      return;
    }

    if (masterFaculty.length === 0) {
      setMessage('Please upload the master faculty file first.');
      event.target.value = '';
      return;
    }

    setLoading(true);
    setAbsentFaculty([]);
    setUnmatchedIds([]);

    try {
      const allAbsentIds = [];
      const processedFileNames = [];
      const fileErrors = [];

      for (const file of selectedFiles) {
        const rows = await readExcelFile(file);

        if (rows.length === 0) {
          fileErrors.push(`${file.name}: No data found`);
          continue;
        }

        const firstValidRow = rows.find((row) => Object.keys(row).length > 0);

        if (!firstValidRow) {
          fileErrors.push(`${file.name}: No valid rows found`);
          continue;
        }

        const staffIdColumn = findColumnName(
          firstValidRow,
          ATTENDANCE_ID_HEADINGS
        );

        const attendanceColumn = findColumnName(
          firstValidRow,
          ATTENDANCE_HEADINGS
        );

        if (!staffIdColumn) {
          fileErrors.push(`${file.name}: Staff ID column was not found`);
          continue;
        }

        if (!attendanceColumn) {
          fileErrors.push(`${file.name}: Attendance column was not found`);
          continue;
        }

        const absentIdsFromFile = rows
          .filter((row) => {
            const attendanceValue = normalizeValue(row[attendanceColumn]);

            return attendanceValue === 'ABSENT';
          })
          .map((row) => normalizeValue(row[staffIdColumn]))
          .filter(Boolean);

        allAbsentIds.push(...absentIdsFromFile);
        processedFileNames.push(file.name);
      }

      if (processedFileNames.length === 0) {
        throw new Error(
          fileErrors.length > 0
            ? fileErrors.join(' | ')
            : 'No valid attendance files were processed.'
        );
      }

      const uniqueAbsentIds = [...new Set(allAbsentIds)];

      const facultyMap = new Map(
        masterFaculty.map((member) => [member.staffId, member])
      );

      const matchedFaculty = [];
      const missingIds = [];

      uniqueAbsentIds.forEach((staffId) => {
        const member = facultyMap.get(staffId);

        if (member) {
          matchedFaculty.push(member);
        } else {
          missingIds.push(staffId);
        }
      });

      const sortedFaculty = matchedFaculty.sort((first, second) => {
        const departmentComparison = (first.department || '').localeCompare(
          second.department || '',
          undefined,
          {
            sensitivity: 'base',
          }
        );

        if (departmentComparison !== 0) {
          return departmentComparison;
        }

        return first.name.localeCompare(second.name, undefined, {
          sensitivity: 'base',
        });
      });

      setAbsentFaculty(sortedFaculty);

      setUnmatchedIds(
        [...new Set(missingIds)].sort((a, b) =>
          a.localeCompare(b, undefined, {
            numeric: true,
            sensitivity: 'base',
          })
        )
      );

      setAttendanceFileNames(processedFileNames);

      if (uniqueAbsentIds.length === 0) {
        setMessage(
          'The selected files do not contain any rows where Attendance is Absent.'
        );
      } else if (sortedFaculty.length === 0) {
        setMessage(
          'Absent Staff IDs were found, but they did not match the master faculty file.'
        );
      } else {
        let resultMessage =
          `${processedFileNames.length} attendance file(s) processed. ` +
          `${uniqueAbsentIds.length} unique absent Staff ID(s) found. ` +
          `${sortedFaculty.length} faculty record(s) matched.`;

        if (fileErrors.length > 0) {
          resultMessage += ` Some files had errors: ${fileErrors.join(' | ')}`;
        }

        setMessage(resultMessage);
      }
    } catch (error) {
      setAttendanceFileNames([]);
      setAbsentFaculty([]);
      setUnmatchedIds([]);
      setMessage(error.message);
    } finally {
      setLoading(false);
      event.target.value = '';
    }
  };

  const copyAbsentFaculty = async () => {
    if (absentFaculty.length === 0) {
      return;
    }

    const header = 'S.No\tFaculty Name\tDepartment';

    const rows = absentFaculty.map(
      (member, index) =>
        `${index + 1}\t${member.name}\t${member.department || '-'}`
    );

    const text = [header, ...rows].join('\n');

    try {
      await navigator.clipboard.writeText(text);
      setMessage('Absent faculty table copied successfully.');
    } catch {
      setMessage(
        'Unable to copy automatically. Please copy the table manually.'
      );
    }
  };

  const downloadAbsentFaculty = () => {
    if (absentFaculty.length === 0) {
      return;
    }

    const worksheetData = absentFaculty.map((member, index) => ({
      'S.No.': index + 1,
      'Faculty Name': member.name,
      Department: member.department || '-',
    }));

    const worksheet = XLSX.utils.json_to_sheet(worksheetData);

    worksheet['!cols'] = [{ width: 10 }, { width: 32 }, { width: 25 }];

    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Absent Faculty');

    XLSX.writeFile(workbook, 'Absent_Faculty_List.xlsx');
  };

  const downloadAbsentFacultyPDF = () => {
    if (absentFaculty.length === 0) {
      return;
    }

    const reportDate = getReportDate();
    const pdf = new jsPDF({
      orientation: 'portrait',
      unit: 'mm',
      format: 'a4',
    });

    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(16);
    pdf.text(
      'RGM College of Engineering and Technology (Autonomous)',
      105,
      18,
      { align: 'center' }
    );

    pdf.setFontSize(13);
    pdf.text('Faculty FRS - Absent List', 105, 27, {
      align: 'center',
    });

    pdf.setFontSize(10);
    pdf.setFont('helvetica', 'normal');
    pdf.text(`Date: ${reportDate.replaceAll('/', '-')}`, 105, 33, {
      align: 'center',
    });

    autoTable(pdf, {
      startY: 39,
      head: [['S.No.', 'Faculty Name', 'Department']],
      body: absentFaculty.map((member, index) => [
        index + 1,
        member.name,
        member.department || '-',
      ]),
      theme: 'grid',
      styles: {
        font: 'helvetica',
        fontSize: 8,
        cellPadding: 1.8,
        valign: 'middle',
      },
      headStyles: {
        fillColor: [13, 90, 167],
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        halign: 'center',
      },
      columnStyles: {
        0: {
          cellWidth: 14,
          halign: 'center',
        },
        1: {
          cellWidth: 75,
        },
        2: {
          cellWidth: 50,
        },
      },
      margin: {
        left: 35.5,
        right: 35.5,
      },
      didDrawPage: () => {
        const pageHeight = pdf.internal.pageSize.getHeight();

        pdf.setFontSize(9);
        pdf.setFont('helvetica', 'normal');

        pdf.text(
          `Page ${pdf.internal.getNumberOfPages()}`,
          105,
          pageHeight - 8,
          { align: 'center' }
        );
      },
    });

    const fileDate = reportDate.replaceAll('/', '-');

    pdf.save(`Faculty_FRS_${fileDate}.pdf`);
  };

  const resetPage = () => {
    setMasterFaculty([]);
    setMasterFileName('');
    setAttendanceFileNames([]);
    setAbsentFaculty([]);
    setUnmatchedIds([]);
    setMessage('Upload the master faculty Excel file first.');
  };

  return (
    <main className="page">
      <section className="container">
        <header className="college-header">
          <div className="college-top">
            <img src="/logo.jpg" alt="RGM Logo" className="college-logo" />

            <div className="college-name">
              <h1>RGM College of Engineering &amp; Technology</h1>
              <h2>(Autonomous)</h2>
            </div>
          </div>

          <div className="department">
            <h1>Face Recognition System (FRS)</h1>

            <p>
              Upload the master faculty file followed by the attendance file.
              Only faculty marked <strong>Absent</strong> will be considered.
            </p>
          </div>
        </header>

        <div className="instructions">
          <strong>Master file:</strong> Staff ID, Faculty Name, Department
          (optional)
          <br />
          <strong>Attendance file:</strong> Sr.No., Staff Id, Staff Name,
          Gender, Enrollment Status, Attendance, In Time
          <br />
          <strong>Selection rule:</strong> Only rows where Attendance is Absent
          are selected.
        </div>

        <div className="upload-grid">
          <section
            className={`upload-card ${
              masterFaculty.length > 0 ? 'completed' : ''
            }`}
          >
            <span className="step-number">1</span>
            <div className="upload-icon">📘</div>
            <h2>Master Faculty File</h2>

            <p>
              Upload the Excel file containing Staff ID and Faculty Name.
              Department is optional.
            </p>

            <label className="file-button">
              Select Master Excel
              <input
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={handleMasterFile}
                disabled={loading}
              />
            </label>

            {masterFileName && (
              <div className="file-details">
                <strong>{masterFileName}</strong>
                <span>{masterFaculty.length} faculty records loaded</span>
              </div>
            )}
          </section>

          <section
            className={`upload-card ${
              attendanceFileNames.length > 0 ? 'completed' : ''
            }`}
          >
            <span className="step-number">2</span>
            <div className="upload-icon">📕</div>
            <h2>Attendance File</h2>

            <p>
              Upload one or more attendance Excel files. Only Absent rows will
              be processed.
            </p>

            <label
              className={`file-button ${
                masterFaculty.length === 0 ? 'disabled' : ''
              }`}
            >
              Select Attendance File
              <input
                type="file"
                accept=".xlsx,.xls,.csv"
                multiple
                onChange={handleAttendanceFiles}
                disabled={loading || masterFaculty.length === 0}
              />
            </label>

            {attendanceFileNames.length > 0 && (
              <div className="file-details">
                <strong>{attendanceFileNames.length} file(s) processed</strong>

                {attendanceFileNames.map((fileName) => (
                  <span key={fileName}>{fileName}</span>
                ))}
              </div>
            )}
          </section>
        </div>

        <div className="status-message">
          {loading ? 'Reading and processing the Excel files...' : message}
        </div>

        {absentFaculty.length > 0 && (
          <section className="results-section">
            <div className="results-header">
              <div>
                <h2>Absent Faculty List</h2>

                <p>
                  Total absent faculty: <strong>{absentFaculty.length}</strong>
                </p>
              </div>

              <div className="action-buttons">
                <button type="button" onClick={copyAbsentFaculty}>
                  Copy
                </button>

                <button type="button" onClick={downloadAbsentFaculty}>
                  Download Excel
                </button>

                <button type="button" onClick={downloadAbsentFacultyPDF}>
                  Download PDF
                </button>
              </div>
            </div>

            <div className="table-wrapper">
              <table className="faculty-table">
                <thead>
                  <tr>
                    <th>S.No.</th>
                    <th>Faculty Name</th>
                    <th>Department</th>
                  </tr>
                </thead>

                <tbody>
                  {absentFaculty.map((member, index) => (
                    <tr key={member.staffId}>
                      <td>{index + 1}</td>
                      <td>{member.name}</td>
                      <td>{member.department || '-'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        )}

        {unmatchedIds.length > 0 && (
          <section className="unmatched-section">
            <h3>Absent Staff IDs not found in the master file</h3>

            <div className="unmatched-list">
              {unmatchedIds.map((id) => (
                <span key={id}>{id}</span>
              ))}
            </div>
          </section>
        )}

        <button className="reset-button" type="button" onClick={resetPage}>
          Reset All
        </button>

        <footer className="app-footer">
          <span>Developed CSE (Data Science)</span>
         
        </footer>
      </section>
    </main>
  );
}

export default App;
