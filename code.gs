

const API_TOKEN = "YOUR_SECRET_TOKEN_HERE";

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents || "{}");

    if (data.token !== API_TOKEN) {
      throw new Error("Unauthorized request.");
    }

    var patientName = String(data.patientName || "").trim();
    var recipient = String(data.email || "").trim();
    var doctor = String(data.doctor || "").trim();
    var appointmentDate = String(data.appointmentDate || "").trim();
    var appointmentTime = String(data.appointmentTime || "").trim();

    if (!patientName || !recipient || !doctor ||
        !appointmentDate || !appointmentTime) {
      throw new Error("Missing appointment information.");
    }

    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(recipient)) {
      throw new Error("Invalid recipient email address.");
    }

    var subject = "Appointment Confirmation";

    var body =
      "Dear " + patientName + ",\n\n" +
      "Your appointment has been scheduled successfully.\n\n" +
      "Doctor: " + doctor + "\n" +
      "Date: " + appointmentDate + "\n" +
      "Time: " + appointmentTime + "\n\n" +
      "Please contact the clinic if you need to make changes.\n\n" +
      "Regards,\nAppointment Team";

    MailApp.sendEmail(recipient, subject, body);

    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        message: "Email submitted successfully."
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: String(error.message || error)
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
