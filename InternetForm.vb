Imports System.Net.Http
Imports System.Threading.Tasks
Imports System.Net
Public Class InternetForm
    Private Sub InternetForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        WebBrowser1.Navigate("https://www.google.com") ' Replace with your preferred webpage
    End Sub

    Private Sub BtnSubmit_Click(sender As Object, e As EventArgs) Handles BtnSubmit.Click
        Dim url As String = "https://docs.google.com/forms/d/e/YOUR_FORM_LINK/viewform"
        WebBrowser1.Navigate(url)
    End Sub

    Private Async Sub btnFetchWeather_Click(sender As Object, e As EventArgs) Handles btnFetchWeather.Click
        Dim client As New HttpClient()

        Dim response As String = Await client.GetStringAsync("https://api.weatherapi.com/v1/current.json?key=f0b0e84ae0b44b5887e85656250403&q=Johannesburg")

        TextBox1.Text = response ' Display response in a TextBox

    End Sub

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted

    End Sub
End Class








