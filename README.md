# Bulk-Email-Sending
Automates sending personalized job application emails with resume attachments directly from Google Sheets using Google Apps Script.
# 📧 Job Application Email Automation (Google Apps Script)

This project automates the process of sending personalized job application emails directly from **Google Sheets** using **Google Apps Script**. Each recipient receives a customized email with your attached resume.

---

## 🚀 Features

- Reads candidate details (Name, Email, Company, Status) from a Google Sheet.
- Skips rows where email status is already marked as **"Sent"**.
- Sends **HTML-formatted personalized emails**.
- Attaches your **resume (PDF)** automatically.
- Updates the **status column** in Google Sheets after sending.

---

## 📂 Project Setup

1. Open [Google Apps Script](https://script.google.com/).
2. Create a **new project**.
3. Copy the contents of [`code.gs`](code.gs) into the script editor.
4. Add the `appsscript.json` configuration file.
5. Replace the `resumeFileId` in the script with your actual **Google Drive file ID**.

---

## 📊 Google Sheet Format

| Name      | Email               | Company      | Status |
|-----------|---------------------|-------------|--------|
| John Doe  | john@example.com    | Google      |        |
| Jane Smith| jane@example.com    | Microsoft   | Sent   |
| Alex Lee  | alex@example.com    | Amazon      |        |

- Leave **Status** blank for pending emails.
- Script will mark it as **Sent** after successful delivery.

---

## 📧 Example Email

- **Subject:** `Application for Business Analyst Opportunities at Google`
- **Body:** Personalized HTML message with introduction, highlights, and LinkedIn link.
- **Attachment:** Resume (PDF)

---

## 🛠 Tech Stack

- Google Apps Script
- Google Sheets
- Gmail API
- Google Drive API

---

## 🔑 Permissions

The script requires the following permissions:
- **Gmail** (to send emails)
- **Google Sheets** (to read/update status)
- **Google Drive** (to fetch resume file)

---


## 💡 Future Improvements

- Add error logging for failed emails.
- Support multiple resume templates.
- Schedule automatic sending via triggers.

---

👩‍💻 Developed by **Vartika**
