function sendJobEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  var resumeFileId = "1-zwo8XC0E_lrXu9NmkjvZd4Ylu1ZD5Kl"; // Replace with your actual file ID
  var file = DriveApp.getFileById(resumeFileId);

  for (var i = 1; i < data.length; i++) {
    var name = data[i][0];
    var email = data[i][1];
    var company = data[i][2];
    var status = data[i][3];

    if (status === "Sent") continue;

    var subject = "Application for Business Analyst Opportunities at " + company;

    var body = `
      <div style="font-family: Arial, sans-serif; line-height:1.6; color:#333;">
        <p>Hi <strong>${name}</strong>,</p>
        <p>I hope this email finds you well. My name is <strong>Vartika</strong>, and I recently completed <strong>BE at Thapar Institute of Engineering & Technology</strong>, specializing in Electronics, Communications.</p>
        <p>I am excited to express my interest in the <strong>Business Analyst internship</strong> at <strong>${company}</strong>. With a blend of technical expertise and business acumen, I am confident in my ability to contribute effectively to your team.</p>
        
        <h3 style="color:#2E86C1;">Key Highlights of My Profile:</h3>
        <ul>
          <li><strong>Business Analysis & Strategy:</strong> Experience in requirements gathering, process mapping, KPI tracking, and competitor research.</li>
          <li><strong>Analytics & Visualization:</strong> Skilled in SQL, Python, Excel, and Power BI; built dashboards improving reporting efficiency by 90%.</li>
          <li><strong>Project & Stakeholder Management:</strong> Partnered with global client teams to streamline workflows and accelerate decision-making.</li>
        </ul>

        <p>I am eager to bring my <strong>data-driven insights</strong> and collaborative mindset to <strong>${company}</strong> and contribute to impactful business outcomes. Please find my resume attached for your review.</p>

        <p>Thank you for your time and consideration. I look forward to your response.</p>

        <p style="margin-top:20px;">
          <strong>Best regards,</strong><br>
          Vartika<br>
          <a href="tel:+918826697597" style="color:#2E86C1;">8826697597</a><br>
          <a href="www.linkedin.com/in/vartika-jaiswal-3279b0188" style="color:#2E86C1;">LinkedIn Profile</a>
        </p>
      </div>
    `;

    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: body,
      attachments: [file.getAs(MimeType.PDF)]
    });

    sheet.getRange(i + 1, 4).setValue("Sent");
  }
}
