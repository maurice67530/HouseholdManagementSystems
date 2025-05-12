<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class In_App_Message
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(In_App_Message))
        Me.btnSendReply = New System.Windows.Forms.Button()
        Me.ListBox2 = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtReply = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.chkShowAll = New System.Windows.Forms.CheckBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnSendReply
        '
        Me.btnSendReply.Location = New System.Drawing.Point(82, 388)
        Me.btnSendReply.Name = "btnSendReply"
        Me.btnSendReply.Size = New System.Drawing.Size(75, 23)
        Me.btnSendReply.TabIndex = 1
        Me.btnSendReply.Text = "Reply"
        Me.btnSendReply.UseVisualStyleBackColor = True
        '
        'ListBox2
        '
        Me.ListBox2.FormattingEnabled = True
        Me.ListBox2.Location = New System.Drawing.Point(82, 281)
        Me.ListBox2.Name = "ListBox2"
        Me.ListBox2.Size = New System.Drawing.Size(430, 95)
        Me.ListBox2.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(79, 81)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(55, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Recipient:"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Azure
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(0, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(577, 40)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Text Messages:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtReply
        '
        Me.txtReply.Location = New System.Drawing.Point(140, 115)
        Me.txtReply.Name = "txtReply"
        Me.txtReply.Size = New System.Drawing.Size(174, 20)
        Me.txtReply.TabIndex = 6
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(79, 118)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(35, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Send:"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(402, 388)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(110, 23)
        Me.Button3.TabIndex = 9
        Me.Button3.Text = "Mark All As Read"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(82, 167)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(430, 95)
        Me.ListBox1.TabIndex = 10
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info
        Me.NotifyIcon1.BalloonTipText = "message from Overdue chores and High Expenses" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.NotifyIcon1.BalloonTipTitle = "House hold Management Notification"
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "House hold Notification"
        Me.NotifyIcon1.Visible = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(83, 265)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(94, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Expense Message"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(83, 151)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(86, 13)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Chores Message"
        '
        'chkShowAll
        '
        Me.chkShowAll.AutoSize = True
        Me.chkShowAll.Location = New System.Drawing.Point(469, 118)
        Me.chkShowAll.Name = "chkShowAll"
        Me.chkShowAll.Size = New System.Drawing.Size(81, 17)
        Me.chkShowAll.TabIndex = 11
        Me.chkShowAll.Text = "CheckBox1"
        Me.chkShowAll.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(140, 78)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(100, 20)
        Me.TextBox2.TabIndex = 14
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Font = New System.Drawing.Font("Lucida Bright", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Red
        Me.Label6.Location = New System.Drawing.Point(-3, 415)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(0, 15)
        Me.Label6.TabIndex = 15
        '
        'In_App_Message
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.ClientSize = New System.Drawing.Size(578, 428)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.chkShowAll)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtReply)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ListBox2)
        Me.Controls.Add(Me.btnSendReply)
        Me.Name = "In_App_Message"
        Me.Text = "In_App_Message"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnSendReply As Button
    Friend WithEvents ListBox2 As ListBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents txtReply As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Button3 As Button
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents NotifyIcon1 As NotifyIcon
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents chkShowAll As CheckBox
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents Label6 As Label
End Class
