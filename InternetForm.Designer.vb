<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InternetForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.WebBrowser1 = New System.Windows.Forms.WebBrowser()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.BtnSubmit = New System.Windows.Forms.Button()
        Me.btnFetchWeather = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'WebBrowser1
        '
        Me.WebBrowser1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WebBrowser1.Location = New System.Drawing.Point(0, 0)
        Me.WebBrowser1.MinimumSize = New System.Drawing.Size(20, 20)
        Me.WebBrowser1.Name = "WebBrowser1"
        Me.WebBrowser1.Size = New System.Drawing.Size(557, 537)
        Me.WebBrowser1.TabIndex = 0
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(39, 12)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(453, 156)
        Me.TextBox1.TabIndex = 1
        '
        'BtnSubmit
        '
        Me.BtnSubmit.Location = New System.Drawing.Point(207, 189)
        Me.BtnSubmit.Name = "BtnSubmit"
        Me.BtnSubmit.Size = New System.Drawing.Size(75, 23)
        Me.BtnSubmit.TabIndex = 2
        Me.BtnSubmit.Text = "SUBMIT"
        Me.BtnSubmit.UseVisualStyleBackColor = True
        '
        'btnFetchWeather
        '
        Me.btnFetchWeather.Location = New System.Drawing.Point(184, 218)
        Me.btnFetchWeather.Name = "btnFetchWeather"
        Me.btnFetchWeather.Size = New System.Drawing.Size(108, 29)
        Me.btnFetchWeather.TabIndex = 3
        Me.btnFetchWeather.Text = "FetchWeather"
        Me.btnFetchWeather.UseVisualStyleBackColor = True
        '
        'InternetForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(557, 537)
        Me.Controls.Add(Me.btnFetchWeather)
        Me.Controls.Add(Me.BtnSubmit)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.WebBrowser1)
        Me.Name = "InternetForm"
        Me.Text = "InternetForm"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents WebBrowser1 As Windows.Forms.WebBrowser
    Friend WithEvents TextBox1 As Windows.Forms.TextBox
    Friend WithEvents BtnSubmit As Windows.Forms.Button
    Friend WithEvents btnFetchWeather As Windows.Forms.Button
End Class
