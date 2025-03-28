
Imports System.Net.Mail
Imports System.Net
Imports System.Windows.Forms

Public Class EmailForm
    Private Sub BtnSend_Click(sender As Object, e As EventArgs) Handles BtnSend.Click

        Try
            Dim smtpClient As New SmtpClient("smtp.gmail.com")
            smtpClient.Port = 587
            smtpClient.Credentials = New NetworkCredential("austinmulalo113@gmail.com", "oqsa qwqa bhjc nzoe")
            smtpClient.EnableSsl = True

            Dim mail As New MailMessage()
            mail.From = New MailAddress("austinmulalo113@gmail.com")
            mail.To.Add(txtTo.Text)
            mail.Subject = txtSubject.Text
            mail.Body = txtBody.Text

            smtpClient.Send(mail)
            MessageBox.Show("Email sent successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub EmailForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class




