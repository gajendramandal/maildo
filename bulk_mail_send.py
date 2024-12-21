import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import pandas as pd
import streamlit as st
import io

st.title("Automatic/Bulk Email Sender")

# Input sender email and password
sender_email = st.text_input("Enter your Email address ")

if sender_email:
    password = st.text_input("Enter your App Password ", type="password")
    st.markdown("[FIND YOUR APP PASSWORD](https://youtu.be/JfTGn2Mm2-Y?si=8GH2Yvl2HQvvATDU)")

    if password:
        # Choose the upload option
        upload_option = st.radio("Choose how to provide recipients", 
                                 ("Upload File (Excel/CSV)", 
                                  "Predefined IIT List",
                                  "Companies HR List",
                                  "Enter Other Emails"))

        recipients = []

        if upload_option == "Upload File (Excel/CSV)":
            uploaded_file = st.file_uploader("Choose your File (Excel/CSV) with 'name' and 'email' column only", 
                                             type=["csv", "xlsx"])

            if uploaded_file is not None:
                try:
                    if uploaded_file.name.endswith('.csv'):
                        df = pd.read_csv(uploaded_file)
                    elif uploaded_file.name.endswith('.xlsx'):
                        df = pd.read_excel(uploaded_file)

                    if "name" in df.columns and "email" in df.columns:
                        recipients = df[['name', 'email']].values.tolist()
                    else:
                        st.error("The file must contain 'name' and 'email' columns.")
                except Exception as e:
                    st.error(f"Error reading file: {e}")

        elif upload_option == "Predefined IIT List":
            df = pd.read_csv('iit_professors_emails.csv')
            recipients = df[['name', 'email']].values.tolist()

        elif upload_option == "Companies HR List":
            df = pd.read_csv('100_companies_list.csv')
            recipients = [(None, email) for email in df['email'].tolist()]

        elif upload_option == "Enter Other Emails":
            other_recipients = st.text_input("Enter the Recipient Emails (separated by commas)")
            if other_recipients:
                recipients = [(email.strip(), email.strip()) for email in other_recipients.split(',')]

        # Input subject and body
        subject = st.text_input("Enter your Subject ")
        st.write("Don't write **DEAR [Name]**, it is automatically generated for each person from uploaded file")
        body = st.text_area("Write the Body", height=300)

        uploaded_file_attach = st.file_uploader("Choose a file to attach", type=["pdf"])

        # Send emails
        if st.button("CLICK TO SEND 👌") and len(recipients) > 0:
            try:
                # Setting up the SMTP server and login
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.starttls()
                server.login(sender_email, password)

                for i, (name, recipient_email) in enumerate(recipients, 1):
                    try:
                        msg = MIMEMultipart()
                        msg['From'] = sender_email
                        msg['To'] = recipient_email
                        msg['Subject'] = subject

                        # Personalize body for different upload options
                        if upload_option == "Upload File (Excel/CSV)":
                            new_body = f"Dear Professor {name},\n\n" + body
                        elif upload_option == "Companies HR List":
                            new_body = f"Dear Hiring Manager,\n\n" + body
                        else:
                            new_body = body

                        msg.attach(MIMEText(new_body, 'plain'))

                        # Handle attachment
                        if uploaded_file_attach is not None:
                            try:
                                pdf_bytes = uploaded_file_attach.getvalue()
                                part = MIMEApplication(pdf_bytes, _subtype="pdf")
                                part.add_header('Content-Disposition', 'attachment', filename=uploaded_file_attach.name)
                                msg.attach(part)
                            except Exception as attach_error:
                                st.warning(f"Could not attach PDF for {recipient_email}: {attach_error}")

                        # Send the email
                        server.sendmail(sender_email, recipient_email, msg.as_string())
                        st.write(f"{i} mail sent to {recipient_email}.")
                    except Exception as email_error:
                        st.error(f"Failed to send email to {recipient_email}. Error: {email_error}")
                        continue

                st.success("Email sending process completed!")
            except Exception as e:
                st.error(f"Error: {e}")
            finally:
                server.quit()
