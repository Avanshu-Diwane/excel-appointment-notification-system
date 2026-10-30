Attribute VB_Name = "Module1"

Option Explicit

Private Const SCRIPT_URL As String = "https://script.google.com/macros/s/AKfycbxUaZjLtuSJwQjaJNlICpMt179ATgfuBmF0Ip1zznx2TZcElP6DdGXS9yUXMvHhSXN4/exec"
Private Const API_TOKEN As String = "ApptNotify_7Kp9vR4xQ2mN8sL6wT5zB3cY"

Public Sub SendAppointmentEmail()

    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim r As Long
    Dim http As Object
    Dim payload As String
    Dim responseText As String
    Dim patientName As String
    Dim recipient As String
    Dim doctor As String
    Dim apptDate As String
    Dim apptTime As String

    Set ws = ThisWorkbook.Worksheets("Appointments")

    If ActiveSheet.Name <> ws.Name Then
        MsgBox "Please select a patient row on the Appointments sheet.", vbExclamation
        Exit Sub
    End If

    r = ActiveCell.Row

    If r < 2 Then
        MsgBox "Please select a patient row, not the header.", vbExclamation
        Exit Sub
    End If

    patientName = Trim(CStr(ws.Cells(r, "B").value))
    recipient = Trim(CStr(ws.Cells(r, "C").value))
    doctor = Trim(CStr(ws.Cells(r, "D").value))

    If Not IsDate(ws.Cells(r, "E").value) And _
   Not IsNumeric(ws.Cells(r, "E").Value2) Then
    MsgBox "Please check the appointment date in column E.", vbExclamation
    Exit Sub
    End If

    If Not IsDate(ws.Cells(r, "F").value) And _
   Not IsNumeric(ws.Cells(r, "F").Value2) Then
    MsgBox "Please check the appointment time in column F.", vbExclamation
    Exit Sub
   End If

    apptDate = Format(ws.Cells(r, "E").value, "dd-mmm-yyyy")
    apptTime = Format(ws.Cells(r, "F").value, "hh:mm AM/PM")

    If patientName = "" Or recipient = "" Or doctor = "" Then
        MsgBox "Patient name, email, or doctor is missing.", vbExclamation
        Exit Sub
    End If

    If InStr(1, SCRIPT_URL, "/exec", vbTextCompare) = 0 _
       Or InStr(1, SCRIPT_URL, "PASTE_", vbTextCompare) > 0 Then
        MsgBox "Please enter your deployed Web App URL in SCRIPT_URL.", vbExclamation
        Exit Sub
    End If

    If InStr(1, API_TOKEN, "PASTE_", vbTextCompare) > 0 _
       Or Len(API_TOKEN) < 20 Then
        MsgBox "Please enter the same secret token used in Apps Script.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Send an appointment confirmation to:" & vbCrLf & _
              recipient & vbCrLf & vbCrLf & _
              "Patient: " & patientName & vbCrLf & _
              "Doctor: " & doctor & vbCrLf & _
              "Date: " & apptDate & vbCrLf & _
              "Time: " & apptTime, _
              vbYesNo + vbQuestion, "Confirm Email") <> vbYes Then
        Exit Sub
    End If

    payload = "{""token"":""" & JsonEscape(API_TOKEN) & """," & _
              """patientName"":""" & JsonEscape(patientName) & """," & _
              """email"":""" & JsonEscape(recipient) & """," & _
              """doctor"":""" & JsonEscape(doctor) & """," & _
              """appointmentDate"":""" & JsonEscape(apptDate) & """," & _
              """appointmentTime"":""" & JsonEscape(apptTime) & """}"

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    http.setTimeouts 10000, 10000, 15000, 30000
    http.Open "POST", SCRIPT_URL, False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send payload

    responseText = http.responseText

    If http.Status >= 200 And http.Status < 300 _
       And InStr(1, responseText, """success"":true", vbTextCompare) > 0 Then

        ws.Cells(r, "G").value = "Sent"
        MsgBox "Email request succeeded for " & patientName & ".", _
               vbInformation, "Appointment Email"

    Else
        ws.Cells(r, "G").value = "Failed"
        MsgBox "The email was not confirmed as successful." & vbCrLf & _
               "HTTP status: " & http.Status & vbCrLf & _
               "Response: " & Left(responseText, 1000), _
               vbExclamation, "Appointment Email"
    End If

CleanExit:
    Set http = Nothing
    Exit Sub

ErrorHandler:
    If Not ws Is Nothing Then
        If r >= 2 Then ws.Cells(r, "G").value = "Failed"
    End If

    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Check the Web App URL, deployment, and internet connection.", _
           vbExclamation, "Appointment Email"
    Resume CleanExit

End Sub

Private Function JsonEscape(ByVal value As String) As String
    value = Replace(value, "\", "\\")
    value = Replace(value, """", "\""")
    value = Replace(value, vbCrLf, "\n")
    value = Replace(value, vbCr, "\n")
    value = Replace(value, vbLf, "\n")
    JsonEscape = value
End Function


