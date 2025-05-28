Imports System.IO
Imports System.Data.OleDb
Imports System.Net.Mail
Imports System.Net

Public Class Settings
    Private Sub Settings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox6.Text = My.Settings.Smtpserver
        TextBox7.Text = My.Settings.SmtpPort.ToString
        TextBox8.Text = My.Settings.EmailFrom
        TextBox9.Text = My.Settings.Password
        TextBox10.Text = My.Settings.RecipientEmail
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        My.Settings.Smtpserver = TextBox6.Text
        My.Settings.SmtpPort = CInt(TextBox7.Text)
        My.Settings.EmailFrom = TextBox8.Text
        My.Settings.Password = TextBox9.Text
        My.Settings.RecipientEmail = TextBox10.Text
        My.Settings.Save()
        MessageBox.Show("Settings saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    Private Sub SendEmailll(subject As String, body As String)
        Try
            Dim smtpClient As New SmtpClient(My.Settings.Smtpserver)
            smtpClient.Port = CInt(My.Settings.SmtpPort)
            smtpClient.Credentials = New NetworkCredential(My.Settings.EmailFrom, My.Settings.Password)
            smtpClient.EnableSsl = True

            Dim mail As New MailMessage()
            mail.From = New MailAddress(My.Settings.EmailFrom)
            mail.To.Add(My.Settings.RecipientEmail) ' Loaded during login
            mail.Subject = subject
            mail.Body = body

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            smtpClient.Send(mail)
            MessageBox.Show("Email sent successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error sending email: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class