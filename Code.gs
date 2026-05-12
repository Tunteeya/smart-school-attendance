// =============================================
// Smart School Attendance & Parent Alert System
// =============================================
// function doGet() {
//   const html = HtmlService.createHtmlOutputFromFile('Index')
//     .setTitle('Smart School Attendance')
//     .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
//     .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');

//   return html;
// }

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Smart School Attendance')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function teacherLogin(username, password) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Teachers");
  const data = sheet.getDataRange().getValues();

  console.log("Trying to login with: " + username + " / " + password);
  console.log("Total rows in Teachers sheet: " + data.length);

  for (let i = 1; i < data.length; i++) {
    const rowUsername = String(data[i][2] || "").trim();
    const rowPassword = String(data[i][3] || "").trim();

    console.log(`Row ${i}: Username="${rowUsername}", Password="${rowPassword}"`);

    if (rowUsername === username && rowPassword === password) {
      return {
        success: true,
        teacherID: data[i][0],
        fullName: data[i][1]
      };
    }
  }
  return { success: false, message: "No matching record found" };
}

// Get classes for a teacher (simple version - all classes for now)
function getClasses() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Classes");
  return sheet.getDataRange().getValues();
}

// Get students in a class
function getStudentsInClass(classID) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
  const data = sheet.getDataRange().getValues();
  const students = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === classID) {
      students.push(data[i]);
    }
  }
  return students;
}


// Mark Attendance + Send Nice Parent Alert
function markAttendance(date, classID, attendanceData) {
  console.log("Received data:", JSON.stringify(attendanceData));

  if (!attendanceData || attendanceData.length === 0) {
    return "Error: No data received";
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attendanceSheet = ss.getSheetByName("Attendance");
  const studentsSheet = ss.getSheetByName("Students");

  let alertCount = 0;

  attendanceData.forEach(record => {
    // Save attendance
    attendanceSheet.appendRow([
      date,
      record.studentID,
      classID,
      record.status,
      record.remarks || "",
      "Teacher"
    ]);

    // Send attractive email if Absent
    if (record.status === "Absent") {
      const allStudents = studentsSheet.getDataRange().getValues();
      const studentRow = allStudents.find(row => String(row[0]) === String(record.studentID));

      if (studentRow && studentRow[4]) {
        const parentEmail = studentRow[4];
        const studentName = studentRow[1];

        const emailBody = `
          <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #ddd; border-radius: 12px; overflow: hidden;">
            <div style="background: linear-gradient(135deg, #1e3c72, #2a5298); color: white; padding: 25px; text-align: center;">
              <h1 style="margin: 0;">⚠️ Absence Alert</h1>
            </div>
            
            <div style="padding: 30px; background: #f9f9f9;">
              <p>Dear Parent/Guardian,</p>
              
              <p>We wish to inform you that your child:</p>
              
              <div style="background: white; padding: 20px; border-radius: 10px; text-align: center; margin: 20px 0;">
                <h2 style="color: #1e3c72; margin: 0;">${studentName}</h2>
                <p style="color: #d32f2f; font-size: 18px; margin: 10px 0;"><strong>Was marked ABSENT</strong></p>
                <p><strong>Date:</strong> ${date}</p>
              </div>
              
              <p>Please contact the school if this was an authorized absence or if you need any assistance.</p>
            </div>
            
            <div style="background: #1e3c72; color: white; text-align: center; padding: 15px; font-size: 14px;">
              Smart School Attendance System<br>
              <small>Keeping parents informed • ${new Date().getFullYear()}</small>
            </div>
          </div>
        `;

        try {
          MailApp.sendEmail({
            to: parentEmail,
            subject: `⚠️ Absence Alert - ${studentName}`,
            htmlBody: emailBody
          });
          alertCount++;
        } catch (e) {
          console.log("Failed to send email to: " + parentEmail);
        }
      }
    }
  });

  return `✅ Attendance saved successfully! ${alertCount} parent alert(s) sent.`;
}


// Test Function - Run this directly
function testViewPast() {
  const testClassID = "C1";   // Change to your actual class ID if different
  const result = viewPastAttendance(testClassID);
  console.log("Test Result for " + testClassID + ":", result);
}

// Generate Attendance Report
function getAttendanceReport(classID) {
  const attendanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance");
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");

  if (!attendanceSheet || !studentsSheet) return { students: [], summary: {} };

  const attendanceData = attendanceSheet.getDataRange().getValues();
  const studentsData = studentsSheet.getDataRange().getValues();

  const studentStats = {};

  // Initialize students in this class
  for (let i = 1; i < studentsData.length; i++) {
    if (studentsData[i][2] === classID) {
      const studentID = studentsData[i][0];
      studentStats[studentID] = {
        name: studentsData[i][1],
        total: 0,
        present: 0
      };
    }
  }

  // Count attendance
  for (let i = 1; i < attendanceData.length; i++) {
    const row = attendanceData[i];
    const studentID = row[1];
    const status = row[3];

    if (studentStats[studentID]) {
      studentStats[studentID].total++;
      if (status === "Present") studentStats[studentID].present++;
    }
  }

  // Calculate percentage
  const report = Object.keys(studentStats).map(id => {
    const s = studentStats[id];
    const percentage = s.total > 0 ? Math.round((s.present / s.total) * 100) : 0;
    return [s.name, s.total, s.present, percentage + "%"];
  });

  return { students: report, totalStudents: Object.keys(studentStats).length };
}


    // ==================== DOWNLOAD AS CSV (Current Attendance) ====================
    function downloadAsCSV() {
      const classID = document.getElementById('classSelect').value || "All";
      const date = new Date().toISOString().split('T')[0];
      const filename = `Attendance_${classID}_${date}.csv`;

      let csvContent = "data:text/csv;charset=utf-8,Date,Student Name,Class,Status,Remarks\n";

      const rows = document.querySelectorAll('#studentsList tr');

      rows.forEach(row => {
        const cells = row.querySelectorAll('td');
        if (cells.length >= 3) {
          const studentName = cells[0].innerText.trim();
          const status = cells[1].innerText.trim();
          const remarks = cells[2] ? cells[2].innerText.trim() : '';
          
          csvContent += `"${date}","${studentName}","${classID}","${status}","${remarks}"\n`;
        }
      });

      if (csvContent.split('\n').length < 3) {
        alert("No data available. Please load students first.");
        return;
      }

      const encodedUri = encodeURI(csvContent);
      const link = document.createElement("a");
      link.setAttribute("href", encodedUri);
      link.setAttribute("download", filename);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      alert("✅ Attendance downloaded successfully!");
    }

    
