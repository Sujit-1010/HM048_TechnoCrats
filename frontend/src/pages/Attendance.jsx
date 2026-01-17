
import React, { useState } from "react";
import { Session, Student, AttendanceRecord, User } from "@/entities/all";
import { SendEmail } from "@/integrations/Core";
import { Card, CardContent } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { ArrowLeft } from "lucide-react"; // Added Calendar and QrCode
import { useNavigate } from "react-router-dom";
import { createPageUrl } from "@/utils";
import { format } from "date-fns";

import SessionSelector from "../components/attendance/SessionSelector";
import AttendanceMarking from "../components/attendance/AttendanceMarking";

export default function AttendancePage() {
  const navigate = useNavigate();
  const [step, setStep] = useState(1);
  const [selectedSession, setSelectedSession] = useState(null);
  const [students, setStudents] = useState([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);

  const handleSessionSelect = async (session) => {
    setIsLoading(true);
    setError(null);
    setSelectedSession(session);
    
    try {
      console.log("=== DEBUG: Session Selection ===");
      console.log("Selected session:", session);
      console.log("Session batch_id:", session.batch_id);
      
      // Primary method: Find students by exact batch_id match
      let studentsForSession = [];
      if (session.batch_id) {
        studentsForSession = await Student.filter({ batch_id: session.batch_id });
      }
      console.log("Students found by batch_id:", studentsForSession.length);
      
      // If no students found, try alternative matching methods
      if (studentsForSession.length === 0) {
        console.log("No students found by batch_id, trying alternative methods...");
        
        // Method 2: Try to find students by college, year, branch, division
        if (session.college_name && session.year && session.branch && session.division) {
          studentsForSession = await Student.filter({
            college_name: session.college_name,
            year: session.year,
            branch: session.branch,
            division: session.division
          });
          console.log("Students found by college/year/branch/division:", studentsForSession.length);
        }
        
        // Method 3: If still no students, try broader search
        if (studentsForSession.length === 0 && session.college_name && session.year) {
          studentsForSession = await Student.filter({
            college_name: session.college_name,
            year: session.year
          });
          console.log("Students found by college/year only:", studentsForSession.length);
        }
      }
      
      console.log("Final students count:", studentsForSession.length);
      if (studentsForSession.length > 0) {
        console.log("Sample student:", studentsForSession[0]);
      }
      
      setStudents(studentsForSession);
      
      // DO NOT UPDATE SESSION STATUS TO ACTIVE HERE - Only update when attendance is actually marked
      
      setStep(2);
    } catch (error) {
      console.error("Error loading students:", error);
      setError(`Error loading students: ${error.message}`);
    }
    
    setIsLoading(false);
  };

  const handleAttendanceComplete = async (attendanceData) => {
    setIsLoading(true);
    setError(null);
    
    try {
      const user = await User.me();
      const presentCount = Object.values(attendanceData).filter(Boolean).length;
      
      const attendanceRecords = students.map(student => ({
        student_id: student.id,
        session_id: selectedSession.id,
        batch_id: selectedSession.batch_id,
        is_present: attendanceData[student.id] || false,
        marked_at: new Date().toISOString(),
        marked_by: user.email
      }));

      await AttendanceRecord.bulkCreate(attendanceRecords);
      
      // NOW update session status to completed
      await Session.update(selectedSession.id, {
        status: "completed",
        present_count: presentCount,
        total_students: students.length
      });
      
      // Generate professional Excel report with styling
      await generateProfessionalReport(attendanceData);
      
      if (selectedSession.tpo_email) {
        try {
          await SendEmail({
            to: selectedSession.tpo_email,
            subject: `Attendance Report - ${selectedSession.session_name}`,
            body: `
Dear TPO,

Attendance has been marked for the following session:

Session Details:
- Session Name: ${selectedSession.session_name}
- Topic: ${selectedSession.topic_taught}
- Faculty: ${selectedSession.faculty_name}
- Date: ${format(new Date(selectedSession.session_date), 'PPP')}
- Time: ${selectedSession.session_time}
- College: ${selectedSession.college_name}
- Year: ${selectedSession.year}
- Branch: ${selectedSession.branch || selectedSession.batch_description}

Attendance Summary:
- Total Students: ${students.length}
- Present: ${presentCount}
- Absent: ${students.length - presentCount}
- Attendance Rate: ${students.length > 0 ? ((presentCount / students.length) * 100).toFixed(1) : 0}%

Best regards,
APT-TECH CRT Solutions
Attendance Management System
            `
          });
        } catch (emailError) {
          console.error("Error sending email:", emailError);
        }
      }
      
      navigate(createPageUrl("Dashboard"));
    } catch (error) {
      console.error("Error completing attendance:", error);
      setError(`Error completing attendance: ${error.message}`);
    }
    
    setIsLoading(false);
  };

  const generateProfessionalReport = async (currentAttendance) => {
    try {
      // Get all sessions for this batch
      const allSessions = await Session.filter({ 
        batch_id: selectedSession.batch_id 
      }, '-session_date', 100);
      
      // Get all attendance records for this batch
      const allAttendanceRecords = await AttendanceRecord.filter({ 
        batch_id: selectedSession.batch_id 
      });
      
      // Build attendance map
      const attendanceMap = {};
      allAttendanceRecords.forEach(record => {
        if (!attendanceMap[record.student_id]) {
          attendanceMap[record.student_id] = {};
        }
        attendanceMap[record.student_id][record.session_id] = record.is_present;
      });
      
      // Add current attendance to the map
      students.forEach(student => {
        if (!attendanceMap[student.id]) {
          attendanceMap[student.id] = {};
        }
        attendanceMap[student.id][selectedSession.id] = currentAttendance[student.id] || false;
      });
      
      // Sort sessions by date
      const sortedSessions = allSessions.sort((a, b) => 
        new Date(a.session_date) - new Date(b.session_date)
      );
      
      // Sort students by roll number
      const sortedStudents = [...students].sort((a, b) => 
        (parseInt(a.roll_no) || 0) - (parseInt(b.roll_no) || 0)
      );
      
      const totalColumns = 3 + sortedSessions.length + 3; // Sr, PRN, Name + Sessions + Present, Total, Attendance %

      let htmlContent = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
        <head>
          <meta charset="UTF-8">
          <!--[if gte mso 9]>
          <xml>
            <x:ExcelWorkbook>
              <x:ExcelWorksheets>
                <x:ExcelWorksheet>
                  <x:Name>Attendance Report</x:Name>
                  <x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>
                </x:ExcelWorksheet>
              </x:ExcelWorksheets>
            </x:ExcelWorkbook>
          </xml>
          <![endif]-->
          <style>
            table { border-collapse: collapse; width: 100%; }
            th, td { border: 1px solid #000; padding: 8px; text-align: left; vertical-align: middle; font-family: Calibri, sans-serif; font-size: 11pt; }
            .header { font-size: 20pt; font-weight: bold; text-align: center; background-color: #DDEBF7; }
            .subheader { font-size: 12pt; font-weight: bold; text-align: center; background-color: #E2EFDA; }
            .section-title { font-size: 12pt; font-weight: bold; background-color: #BFBFBF; color: #FFFFFF; text-align: left; }
            .bold { font-weight: bold; }
            .present { background-color: #C6EFCE; color: #006100; text-align: center; font-weight: bold; }
            .absent { background-color: #FFC7CE; color: #9C0006; text-align: center; font-weight: bold; }
            .summary-header { background-color: #F8CBAD; font-weight: bold; }
            .info-cell { background-color: #FCE4D6; }
          </style>
        </head>
        <body>
      `;
      
      htmlContent += `
        <table>
          <tr><td colspan="${totalColumns}" class="header">APT-TECH Professional Training Attendance Report</td></tr>
          <tr><td colspan="${totalColumns}" class="subheader">${selectedSession.college_name}</td></tr>
          <tr><td colspan="${totalColumns}" class="subheader">Class: ${selectedSession.batch_description || selectedSession.branch} (${selectedSession.year})</td></tr>
          <tr><td colspan="${totalColumns}"></td></tr> <!-- Spacer row -->
        </table>
      `;

      htmlContent += `
        <table>
          <tr><td colspan="6" class="section-title">SESSION DETAILS</td></tr>
          <tr class="bold"><td>Date</td><td>Session Name</td><td>Topic Taught</td><td>Faculty</td><td>Time</td><td>Duration</td></tr>
      `;
      sortedSessions.forEach(sess => {
        htmlContent += `
          <tr>
            <td>${format(new Date(sess.session_date), 'dd/MM/yy')}</td>
            <td>${sess.session_name || ''}</td>
            <td>${sess.topic_taught || ''}</td>
            <td>${sess.faculty_name || ''}</td>
            <td>${sess.session_time || ''}</td>
            <td>${sess.duration || ''}</td>
          </tr>
        `;
      });
      htmlContent += `</table><br/>`;

      htmlContent += `
        <table>
          <tr><td colspan="${totalColumns}" class="section-title">ATTENDANCE RECORD</td></tr>
          <tr class="bold"><td>Sr.</td><td>PRN</td><td>Student Name</td>
      `;
      sortedSessions.forEach(sess => {
        htmlContent += `<td>${format(new Date(sess.session_date), 'dd/MM/yy')}</td>`;
      });
      htmlContent += `<td>Present</td><td>Total</td><td>Attendance %</td></tr>`;
      
      sortedStudents.forEach((student, index) => {
        let presentCount = 0;
        let attendanceRow = '';
        sortedSessions.forEach(sess => {
          const isPresent = attendanceMap[student.id] && attendanceMap[student.id][sess.id];
          if (isPresent) presentCount++;
          attendanceRow += `<td class="${isPresent ? 'present' : 'absent'}">${isPresent ? 'P' : 'A'}</td>`;
        });

        const totalSess = sortedSessions.length;
        const percentage = totalSess > 0 ? Math.round((presentCount / totalSess) * 100) : 0;
        
        htmlContent += `
          <tr>
            <td>${index + 1}</td>
            <td>${student.prn || ''}</td>
            <td class="bold">${student.name || ''}</td>
            ${attendanceRow}
            <td class="bold">${presentCount}</td>
            <td class="bold">${totalSess}</td>
            <td class="bold">${percentage}%</td>
          </tr>
        `;
      });
      htmlContent += `</table><br/>`;

      const totalStudents = sortedStudents.length;
      const totalSessions = sortedSessions.length;
      const totalActualAttendance = sortedStudents.reduce((sum, student) => {
        return sum + sortedSessions.reduce((sessSum, sess) => {
          return sessSum + (attendanceMap[student.id] && attendanceMap[student.id][sess.id] ? 1 : 0);
        }, 0);
      }, 0);
      const totalPossibleAttendance = totalStudents * totalSessions;
      const overallAttendanceRate = totalPossibleAttendance > 0 ? Math.round((totalActualAttendance / totalPossibleAttendance) * 100) : 0;

      htmlContent += `
        <table>
          <tr><td colspan="2" class="section-title">SUMMARY</td></tr>
          <tr><td class="bold summary-header">Total Students:</td><td class="info-cell">${totalStudents}</td></tr>
          <tr><td class="bold summary-header">Total Sessions Completed:</td><td class="info-cell">${totalSessions}</td></tr>
          <tr><td class="bold summary-header">Overall Attendance Rate:</td><td class="info-cell">${overallAttendanceRate}%</td></tr>
          <tr><td class="bold summary-header">Report Generated:</td><td class="info-cell">${format(new Date(), 'dd/MM/yyyy HH:mm')}</td></tr>
          <tr><td class="bold summary-header">Generated By:</td><td class="info-cell">APT-TECH Attendance System</td></tr>
          <tr><td class="bold summary-header">Contact:</td><td class="info-cell">sanirkittur@apt-techsolution.com</td></tr>
          <tr><td class="bold summary-header">Website:</td><td class="info-cell">https://apt-techsolution.com</td></tr>
        </table>
      `;

      htmlContent += `</body></html>`;

      const blob = new Blob([htmlContent], { type: 'application/vnd.ms-excel;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      
      const fileName = `${selectedSession.college_name}_${selectedSession.batch_description || selectedSession.branch}_Report.xls`.replace(/[\s/\\?%*:|"<>]/g, '_');
      
      link.download = fileName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
      
    } catch (error) {
      console.error("Error generating professional report:", error);
    }
  };

  const handleBackToSessions = () => {
    // Reset to initial state without changing any session status
    setStep(1);
    setSelectedSession(null);
    setStudents([]);
    setError(null);
  };

  return (
    <div className="min-h-screen bg-gray-50 p-4 md:p-6">
      <div className="max-w-7xl mx-auto">
        <div className="flex items-center gap-4 mb-6">
          <Button
            variant="outline"
            size="icon"
            onClick={() => navigate(createPageUrl("Dashboard"))}
            className="hover:bg-gray-100"
          >
            <ArrowLeft className="w-5 h-5" />
          </Button>
          <div>
            <h1 className="text-2xl font-bold text-gray-900">📋 Take Attendance</h1>
            <p className="text-gray-600 mt-1">Select a session and choose your attendance method</p>
          </div>
        </div>

        {error && (
          <div className="bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded mb-4" role="alert">
            <strong className="font-bold">Error! </strong>
            <span className="block sm:inline">{error}</span>
          </div>
        )}

        <div className="space-y-6">
          {step === 1 && (
            <SessionSelector 
              onSessionSelect={handleSessionSelect}
              onQRSessionSelect={(session) => navigate(createPageUrl(`QRAttendance?session_id=${session.id}`))}
              isLoading={isLoading}
            />
          )}

          {step === 2 && selectedSession && (
            <>
              {students.length === 0 ? (
                <Card className="bg-white border border-gray-200">
                  <CardContent className="p-8 text-center">
                    <h3 className="text-lg font-semibold text-gray-900 mb-2">📭 No Students Found</h3>
                    <p className="text-gray-600 mb-4">
                      No students found for this session's batch. Please ensure students are uploaded with matching details.
                    </p>
                    <Button 
                      onClick={() => navigate(createPageUrl("Upload"))}
                      className="bg-blue-600 hover:bg-blue-700"
                    >
                      📤 Upload Students
                    </Button>
                  </CardContent>
                </Card>
              ) : (
                <AttendanceMarking 
                  session={selectedSession}
                  students={students}
                  onComplete={handleAttendanceComplete}
                  onBack={handleBackToSessions}
                  isLoading={isLoading}
                />
              )}
            </>
          )}
        </div>
      </div>
    </div>
  );
}
