Imports Microsoft.VisualBasic.ApplicationServices

Namespace My
    Partial Friend Class MyApplication
        Inherits WindowsFormsApplicationBase

        Public Sub New()
            MyBase.New()
            Me.IsSingleInstance = True
            Me.EnableVisualStyles = True
            Me.ShutdownStyle = ShutdownMode.AfterMainFormCloses
        End Sub

        Protected Overrides Sub OnCreateMainForm()
            Me.MainForm = New Login()
        End Sub
    End Class
End Namespace