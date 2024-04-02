import win32com.client as win32

email_outlook = win32.Dispatch('outlook.application')
email = email_outlook.CreateItem(0)
email.To = 'savio.vinw@gmail.com'
email.Subject = 'email vindo do outlook'
email.Body = 'texto do email'

attachment = r'C:\Users\Sony\Documents\Curriculo\atuais\SAVIO_SOUSA_EstagioSuporteComFoto.pdf'

email.Attachments.Add(attachment)

email.Send()