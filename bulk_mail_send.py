import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import pandas as pd
import streamlit as st
import io
# from frontend import render_frontend
st.title("Automatic/Bulk Email sender")
# render_frontend()
sender_email = st.text_input("Enter your Email address ")

if sender_email:
    password = st.text_input("Enter your App Password ", type="password")
    st.markdown("[FIND YOUR APP PASSWORD](https://youtu.be/JfTGn2Mm2-Y?si=8GH2Yvl2HQvvATDU )")

    if password:
        # Choose file upload or prelist option
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
        
        # Subject and body inputs
        subject = st.text_input("Enter your Subject ")
        st.write("Don't write **DEAR [Name]**, it is automatically generated for each person from uploaded file")
        body = st.text_area("Write the Body", height=300)
        
        uploaded_file_attach = st.file_uploader("Choose a file to attach", type=["pdf"])

        if st.button("CLICK TO SEND 👌") and len(recipients) > 0:
            try:
                # Setting up the SMTP server and login
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.starttls()  
                server.login(sender_email, password)

                # Send email
                for i, (name, recipient_email) in enumerate(recipients, 1):
                    msg = MIMEMultipart()
                    msg['From'] = sender_email
                    msg['To'] = recipient_email
                    msg['Subject'] = subject
                    
                    # Personalize body for Excel/CSV uploads
                    if upload_option == "Upload File (Excel/CSV)":
                        old_body = body
                        # new_body = body
                        new_body = f"Dear {name},\n\n" + body
                        
                    elif upload_option == "Companies HR List":
                        old_body = body
                        # new_body = body
                        new_body = f"Dear Hiring Manager,\n\n" + body
                    
                    else:
                        new_body=body
                                     
                    msg.attach(MIMEText(new_body, 'plain'))

                    # Improved PDF attachment handling
                    if uploaded_file_attach is not None:
                        try:
                            # Read the file into memory
                            pdf_bytes = uploaded_file_attach.getvalue()
                            
                            # Create a new file-like object in memory
                            pdf_file = io.BytesIO(pdf_bytes)
                            
                            # Create PDF attachment
                            part = MIMEApplication(pdf_file.read(), _subtype="pdf")
                            part.add_header('Content-Disposition', 'attachment', filename=uploaded_file_attach.name)
                            msg.attach(part)
                            
                            # Reset file pointer
                            pdf_file.seek(0)
                        except Exception as attach_error:
                            st.warning(f"Could not attach PDF for {name}: {attach_error}")
                    
                    # Send the email
                    server.sendmail(sender_email, recipient_email, msg.as_string())
                    
                    # Provide progress feedback
                    if upload_option == "Upload File (Excel/CSV)":
                        st.write(f"{i} mail sent to {name}.")
                        new_body = old_body
                    elif upload_option == "Companies HR List":
                        st.write(f"{i} mail sent.")
                        new_body = old_body
                    else:
                        st.write(f"{i} mail sent.")  
                        
                st.success("All Emails sent successfully!")
            except Exception as e:
                st.error(f"Error: {e}")
            finally:
                server.quit()
