# Bulk Email Sender

A user-friendly Streamlit app designed to send personalized bulk emails effortlessly. Whether it's uploading recipient lists or using predefined emails, this app simplifies bulk mailing for students, professionals, and researchers. 🎉

## Live App
Try it out here: [Bulk Email Sender](https://bulkymail.streamlit.app) 🚀

---

## Features
- **Multiple Recipient Options:**
  - Upload recipient lists via Excel/CSV files.
  - Use a predefined list of IIT professor emails.
  - Use a Companies HR List
  - Manually input email addresses.
- **Email Personalization:**
  - Automatically adds personalized salutations for recipients.
- **Attachment Support:**
  - Attach PDF files to your emails.
- **Streamlined Workflow:**
  - Clean and intuitive interface for sending emails in bulk.

---

## How to Use
1. **Log in with your Gmail account:**
   - Enter your Gmail address.
   - Use your [App Password](https://www.youtube.com/watch?v=N_J3HCATA1c) for secure login.

2. **Choose Recipients:**
   - Upload an Excel/CSV file containing `name` and `email` columns.
   - Select the predefined IIT professor list.
   - Enter recipient emails manually, separated by commas.

3. **Compose Your Email:**
   - Write the subject line and body. (The app automatically adds "Dear [Name]" for personalized emails.)
   - Optionally, attach a PDF file.

4. **Send Emails:**
   - Click the "CLICK TO SEND 👌" button to start sending your emails.

---

## Installation (Optional for Local Use)
To run this app locally:
1. Clone the repository:
   ```bash
   git clone https://github.com/your-repo/bulk-email-sender.git
   cd bulk-email-sender
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the app:
   ```bash
   streamlit run bulk_mail_send.py
   ```

---

## Technical Stack
- **Frontend:** [Streamlit](https://streamlit.io)
- **Backend:** Python (SMTP protocol for email sending)
- **Data Handling:** Pandas for processing Excel/CSV files

---

## Security Notice
- **Your email credentials are not stored** by this app.
- Use a Gmail app-specific password for enhanced security.
- Always keep sensitive information safe when sharing files or credentials.

---

## Acknowledgments
Special thanks to Baibhav for providing the IIT email list used in this project.

---

Enjoy seamless bulk emailing with this app! Feel free to contribute or suggest improvements via GitHub. 🙂

