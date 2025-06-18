# 📧 Email Automation System (Google Apps Script)

This is a professional-grade email automation system built using Google Sheets + Apps Script.

## ✅ Features

- HTML-based email templates
- Test mode (no actual emails sent)
- Auto logging (date, name, subject, status)
- Status toggle system (ON/OFF)
- Column, Line, and Pie charts for sent email analytics
- Optional file attachments

## 📊 Dashboard Preview

- Emails per day (Column + Line Chart)
- Status breakdown (Pie Chart)

## 🔒 Test Mode

Emails can be previewed without actually sending using `TEST_MODE = true`.

## 🔗 Usage

1. Copy script to Apps Script Editor.
2. Link Google Sheet with appropriate headers:
   - Name | Email | Subject | Message | Status | File ID | Last Subject | Last Message | Toggle | Last Sent | Template Name

3. Add your HTML email templates in Apps Script under "HTML".

4. Run `previewAndSendEmails` or use the `📧 Email Actions` menu.


