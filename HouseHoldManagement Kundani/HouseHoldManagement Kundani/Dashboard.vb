Imports System.IO
Imports System.Data.OleDb
Imports System.Net.Mail
Imports System.Net
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.DataVisualization.Charting
Public Class Dashboard

    ' Daily tips list
    Dim tips As New List(Of String) From {
    "Stay organized and save time.",
    "Complete one task at a time.",
    "Keep your groceries fresh and updated.",
    "A clean home is a happy home.",
    "Plan your meals to avoid waste.",
    "Track your expenses to stay on budget.",
    "Small tasks done daily keep chores away.",
    "Involve everyone – teamwork works best!"
}
    ' Set your design size (the size you built in designer)
    Private originalWidth As Integer
    Private originalHeight As Integer
    Public Shared LoggedInUser As String
    Private Sub Dashboard_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Show total users who have logged in
        Label39.Text = "Users: " & GetTotalUsersLoggedIn().ToString()
        CheckDatabaseConnection()


        lblbadge.Region = New Region(New Drawing2D.GraphicsPath())
        Dim gp As New Drawing.Drawing2D.GraphicsPath()
        gp.AddEllipse(0, 0, lblbadge.Width, lblbadge.Height)
        lblbadge.Region = New Region(gp)
        StyleBadge()
        In_App_Message.NotifyIcon1.ShowBalloonTip(5000) ' 5 seconds


        ' Save the form's original design size
        originalWidth = Me.Width
        originalHeight = Me.Height

        '' Get screen size excluding the taskbar
        'Dim workingWidth As Integer = Screen.PrimaryScreen.WorkingArea.Width
        'Dim workingHeight As Integer = Screen.PrimaryScreen.WorkingArea.Height

        '' Calculate scale
        'Dim scaleX As Single = workingWidth / originalWidth
        'Dim scaleY As Single = workingHeight / originalHeight

        '' Resize the form to fit working area
        'Me.FormBorderStyle = FormBorderStyle.None
        'Me.Bounds = Screen.PrimaryScreen.WorkingArea
        'Me.AutoScaleMode = AutoScaleMode.None
        'Application.DoEvents()

        '' Resize controls
        'ResizeControls(Me, scaleX, scaleY)

        '' Allow ESC to close while testing
        'Me.KeyPreview = True
        ''''''''

        ' Save the form's original design size
        originalWidth = Me.Width
        originalHeight = Me.Height

        ' Get screen size excluding the taskbar
        Dim workingWidth As Integer = Screen.PrimaryScreen.WorkingArea.Width
        Dim workingHeight As Integer = Screen.PrimaryScreen.WorkingArea.Height

        ' Calculate scale
        Dim scaleX As Single = workingWidth / originalWidth
        Dim scaleY As Single = workingHeight / originalHeight

        ' Resize the form to fit working area
        Me.FormBorderStyle = FormBorderStyle.None
        Me.Bounds = Screen.PrimaryScreen.WorkingArea
        Me.AutoScaleMode = AutoScaleMode.None
        Application.DoEvents()

        ' Resize controls
        ResizeControls(Me, scaleX, scaleY)

        ' Allow ESC to close while testing
        Me.KeyPreview = True

        UpdateNotificationCount()
        LoadChoresStatus()


        LoadChoresStatus()
        LoadChart()
        LoadUpcomingMeals()
        LoadExpiringGroceries()
        PopulateListboxFromTasks(Tasks)
        ShowChoreStatusPieChart()
        LoadRecentPhotos()
        DisplayPhoto()
        SetupTimer()
        'StartPhotoSlideshow()
        LoadRecentPhotos()      ' Load from database
        DisplayPhoto()          ' Display the first photo
        SetupTimer()
        LoadMemberChoreTaskSummary()
        'LoadUnassignedPeople()
        LoadAssignmentSummary()
        LoadUpcomingDates()
        LoadUpcomingBirthdays()
        LoadRecentGroceries()


        ' Show current date, month, and time
        Label17.Text = DateTime.Now.ToString("dddd, dd MMMM yyyy - hh:mm tt")
        'Label17.Location = New Point(10, 10) 'Adjust position as needed
        Label17.AutoSize = True
        Label17.Font = New Font("Arial", 10, FontStyle.Bold)
        Label17.ForeColor = Color.Blue


        Dim hour As Integer = DateTime.Now.Hour
        If hour < 12 Then
            Label18.Text = "Good morning , Have a great day! 😊"
        ElseIf hour < 18 Then
            Label18.Text = "Good afternoon , Have a great day! 😊"
        Else
            Label18.Text = "Good evening ,  Have a great day! 😊"
        End If

        ''Dim user As String = LoggedInUser
        ''Dim hour As Integer = DateTime.Now.Hour
        ''Dim greeting As String = ""

        ''If hour >= 5 AndAlso hour < 12 Then
        ''    greeting = "Good Morning"
        ''ElseIf hour >= 12 AndAlso hour < 17 Then
        ''    greeting = "Good Afternoon"
        ''ElseIf hour >= 17 AndAlso hour < 21 Then
        ''    greeting = "Good Evening"
        ''Else
        ''    greeting = "Good Night"
        ''End If

        'Label18.Text = $"{greeting}, {user}!"
        'Label18.ForeColor = Color.DarkBlue
        'Label18.Font = New Font("Segoe UI", 14, FontStyle.Bold)

        'CheckExpense()




        'Show Random tip
        Dim rnd As New Random()
        Dim index As Integer = rnd.Next(tips.Count)
        Label21.Text = "Tip of the Day: " & tips(index)


        ToolTip1.SetToolTip(Button7, "Task")
        ToolTip1.SetToolTip(Button15, "Inventory")
        ToolTip1.SetToolTip(Button12, "Expense")
        ToolTip1.SetToolTip(Button11, "Chores")
        ToolTip1.SetToolTip(Button14, "MealPlan")
        ToolTip1.SetToolTip(Button13, "GroceryItem")
        ToolTip1.SetToolTip(Button16, "Notification")
        ToolTip1.SetToolTip(Button5, "Personel")
        ToolTip1.SetToolTip(Button8, "PhotoGallery")
        ToolTip1.SetToolTip(Button6, "Family Event")
        ToolTip1.SetToolTip(btnInAppMessages, "Notifications Status")
        ToolTip1.SetToolTip(Button17, "Budget")

        Timer1.Interval = 100
        Timer1.Start()
        If photoList.Count > 0 Then
            DisplayPhoto() ' Show the first photo
            SetupTimer() ' Start timer for slideshow
        End If


        Timer2.Interval = 200
        Timer2.Start()

        Timer3.Interval = 200
        Timer3.Start()

        LoadFamilyScheduleAlerts()

        ShowInternetSpeed()


        If currentUser = "Admin" Then
            ' Admin has full access

            Button15.Enabled = True
            Button11.Enabled = True
            Button13.Enabled = True
            Button14.Enabled = True
            Button7.Enabled = True
            Button8.Enabled = True
            Button12.Enabled = True
            Button17.Enabled = True
            Button6.Enabled = True
            Button16.Enabled = True
            Button5.Enabled = True
            btnInAppMessages.Enabled = True

        ElseIf currentUser = "Member" Then
            ' Members have limited access

            Button11.Enabled = True
            Button5.Enabled = True
            Button7.Enabled = True
            Button6.Enabled = True
            Button8.Enabled = True

            Button15.Enabled = False
            Button13.Enabled = False
            Button14.Enabled = False
            Button12.Enabled = False
            Button17.Enabled = False
            Button16.Enabled = False
            btnInAppMessages.Enabled = False

        ElseIf currentUser = "Chef" Then

            Button15.Enabled = True
            Button13.Enabled = True
            Button14.Enabled = True
            btnInAppMessages.Enabled = True

            Button7.Enabled = False
            Button2.Enabled = False
            Button8.Enabled = False
            Button6.Enabled = False
            Button1.Enabled = False
            Button10.Enabled = False


        ElseIf currentUser = "Finance" Then

            Button12.Enabled = True
            Button17.Enabled = True

            Button3.Enabled = False
            Button4.Enabled = False
            Button7.Enabled = False
            Button5.Enabled = False
            Button2.Enabled = False
            Button8.Enabled = False
            Button6.Enabled = False
            Button10.Enabled = False

        End If


    End Sub

    Private Sub ResizeControls(parent As Control, scaleX As Single, scaleY As Single)
        'For Each ctrl As Control In parent.Controls
        '    ctrl.Left = CInt(ctrl.Left * scaleX)
        '    ctrl.Top = CInt(ctrl.Top * scaleY)
        '    ctrl.Width = CInt(ctrl.Width * scaleX)
        '    ctrl.Height = CInt(ctrl.Height * scaleY)

        '    Dim fontScale As Single = (scaleX + scaleY) / 2
        '    ctrl.Font = New Font(ctrl.Font.FontFamily, ctrl.Font.Size * fontScale, ctrl.Font.Style)

        '    If ctrl.HasChildren Then
        '        ResizeControls(ctrl, scaleX, scaleY)
        '    End If
        'Next


        For Each ctrl As Control In parent.Controls
            ctrl.Left = CInt(ctrl.Left * scaleX)
            ctrl.Top = CInt(ctrl.Top * scaleY)
            ctrl.Width = CInt(ctrl.Width * scaleX)
            ctrl.Height = CInt(ctrl.Height * scaleY)

            Dim fontScale As Single = (scaleX + scaleY) / 2
            ctrl.Font = New Font(ctrl.Font.FontFamily, ctrl.Font.Size * fontScale, ctrl.Font.Style)

            If ctrl.HasChildren Then
                ResizeControls(ctrl, scaleX, scaleY)
            End If
        Next

    End Sub

    ' ESC key to close (for testing)
    Private Sub MainForm_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Function GetTotalUsersLoggedIn() As Integer

        Dim total As Integer = 0

        Dim query As String = "SELECT COUNT(*) FROM (SELECT UserName FROM Users GROUP BY UserName)"

        Using conn As New OleDbConnection(connectionString)

            Using cmd As New OleDbCommand(query, conn)

                conn.Open()

                total = CInt(cmd.ExecuteScalar())

            End Using

        End Using

        Return total

    End Function


    Private Sub LoadChoresStatus()
        Dim completed As Integer = 0, inProgress As Integer = 0, notStarted As Integer = 0
        Dim query As String = "SELECT Status, COUNT(*) FROM Chores GROUP BY Status"

        Using conn As New OleDbConnection(connectionString),
          cmd As New OleDbCommand(query, conn)

            conn.Open()
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim status As String = reader("Status").ToString().Trim().ToLower()
                    Select Case status
                        Case "completed"
                            completed = Convert.ToInt32(reader(1))
                        Case "in progress"
                            inProgress = Convert.ToInt32(reader(1))
                        Case "not started"
                            notStarted = Convert.ToInt32(reader(1))
                    End Select
                End While
            End Using
        End Using

        Label15.Text = $"Chores: - Completed: {completed} - In Progress: {inProgress} - Not Started: {notStarted}"
    End Sub


    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        Task_Management.ShowDialog()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Grocery_Items.ShowDialog()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        chores.ShowDialog()
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Inventory.ShowDialog()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        MealPlan.ShowDialog()
    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click
        PhotoGallery.ShowDialog()
    End Sub


    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Expense.ShowDialog()
    End Sub

    Private Sub Button6_Click_1(sender As Object, e As EventArgs) Handles Button6.Click
        Family_Schedule.ShowDialog()
    End Sub

    Private Sub Button5_Click_1(sender As Object, e As EventArgs) Handles Button5.Click
        Personnel.ShowDialog()
    End Sub


    Private Function CheckExpense() As Boolean
        Dim BudgetLimit As Decimal = 700
        Dim TotalExpense As Decimal = 0

        Try
            Dim Conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
            Conn.Open()
            Dim cmd As New OleDbCommand("SELECT Amount From Expense", Conn)
            Dim Reader As OleDbDataReader = cmd.ExecuteReader

            While Reader.Read()
                TotalExpense += Convert.ToDecimal(Reader("Amount"))
            End While

            Reader.Close()
            Conn.Close()

            If TotalExpense >= (BudgetLimit * 0.8D) Then
                ' More than or equal to 80% – Alert
                MessageBox.Show("Alert! You have used more than 80% of your budget", "Budget Alert", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                Dim messageBody As String = $"Alert! Budget Alert:{vbCrLf}{vbCrLf}Total Expenses: {TotalExpense}{vbCrLf}Budget Limit: {BudgetLimit}"
                SendEmail("nethonondamudzunga45@gmail.com", "Budget Alert", messageBody)
                MessageBox.Show("Budget Alert Sent Successfully!", "Budget Alert", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                ' Less than 80% – Budget under control
                MessageBox.Show("Your budget is under control.", "Budget Status", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Dim messageBody As String = $"Budget Status: Under Control{vbCrLf}{vbCrLf}Total Expenses: {TotalExpense}{vbCrLf}Budget Limit: {BudgetLimit}"
                SendEmail("nethonondamudzunga45@gmail.com", "Budget Status", messageBody)
                MessageBox.Show("Budget Status Email Sent Successfully!", "Budget Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show("An error occurred while checking the budget.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return False
    End Function

    ' Function to check expired groceries from Inventory table

    Private Function CheckOverdueChores() As Boolean

        Dim overdueChore As String = ""

        ' Modify the query to retrieve expired groceries

        Dim query As String = "SELECT Status, DueDate FROM Chores" ' Adjust query based on your table

        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)

            Try

                conn.Open()

                Using command As New OleDbCommand(query, conn)

                    Using reader As OleDbDataReader = command.ExecuteReader()

                        While reader.Read()

                            Dim Status As String = reader("Status").ToString()

                            Dim DueDate As Date = Convert.ToDateTime(reader("DueDate"))

                            overdueChore &= $"{Status} Due on {DueDate.ToShortDateString()}" & vbCrLf

                        End While

                    End Using

                End Using

            Catch ex As Exception

                'MessageBox.Show("Error retrieving expired overdueChore: " & ex.Message)

                Return False
                conn.Close()
            End Try

        End Using

        If Not String.IsNullOrEmpty(overdueChore) Then

            ' Display MessageBox showing expired items

            MessageBox.Show($"overdueChores:{vbCrLf}{vbCrLf}{overdueChore}", "overdueChore", MessageBoxButtons.OK, MessageBoxIcon.Warning)


            ' Send the email with the expired items

            Dim messageBody As String = $"Alert! overdueChore:{vbCrLf}{vbCrLf}{overdueChore}"

            SendEmail("nethonondamudzunga45@gmail.com", "overdueChore Alert", messageBody)

            ' Notify that the email was sent

            MessageBox.Show("overdueChore Alert Sent Successfully!", "overdueChore", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If


        Return False

    End Function
    ' Function to check expired groceries from Inventory table

    Private Function CheckExpiredGroceries() As Boolean

        Dim expiredGroceries As String = ""

        ' Modify the query to retrieve expired groceries

        Dim query As String = "SELECT ItemName, ExpiryDate FROM Inventory" ' Adjust query based on your table

        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)

            Try

                conn.Open()

                Using command As New OleDbCommand(query, conn)

                    Using reader As OleDbDataReader = command.ExecuteReader()

                        While reader.Read()

                            Dim itemName As String = reader("ItemName").ToString()

                            Dim expiryDate As Date = Convert.ToDateTime(reader("ExpiryDate"))

                            expiredGroceries &= $"{itemName} expired on {expiryDate.ToShortDateString()}" & vbCrLf

                        End While

                    End Using

                End Using

            Catch ex As Exception

                MessageBox.Show("Error retrieving expired groceries: " & ex.Message)

                Return False
                conn.Close()

            End Try

        End Using

        If Not String.IsNullOrEmpty(expiredGroceries) Then

            ' Display MessageBox showing expired items

            MessageBox.Show($"The following grocery items have expired:{vbCrLf}{vbCrLf}{expiredGroceries}", "Expired Groceries", MessageBoxButtons.OK, MessageBoxIcon.Warning)


            ' Send the email with the expired items

            Dim messageBody As String = $"Alert! The following grocery items have expired:{vbCrLf}{vbCrLf}{expiredGroceries}"

            SendEmail("nethonondamudzunga45@gmail.com", "Grocery Expiry Alert", messageBody)

            ' Notify that the email was sent

            MessageBox.Show("Expired Groceries Alert Sent Successfully!", "Expired Groceries", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If


        Return False

    End Function

    ' Function to send email

    Private Sub SendEmail(recipient As String, subject As String, messageBody As String)

        Try

            ' Configure SMTP client

            Dim smtpClient As New SmtpClient(My.Settings.Smtpserver) With {.Port = 587, .EnableSsl = True, .Credentials = New NetworkCredential("nethonondamudzunga45@gmail.com", "slwo xavj lool amzu")}


            ' Create the email message

            Dim mailMessage As New MailMessage() With {.From = New MailAddress("nethonondamudzunga45@gmail.com"), .Subject = subject, .Body = messageBody}



            ' Add recipient

            mailMessage.To.Add(recipient)

            ' Send the email

            smtpClient.Send(mailMessage)

        Catch ex As Exception

            MessageBox.Show("Error sending email: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Timer1.Stop()

        CheckExpiredGroceries()

        CheckOverdueChores()

        CheckExpense()

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim budgetLimit As Double = 700
    Dim blinkState As Boolean = True

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        Dim totalExpenses = GetTotalExpenses()
        Dim totalIncome = GetTotalIncome()
        Dim remaining = budgetLimit - totalExpenses
        Dim usedPercent As Double = (totalExpenses / budgetLimit) * 100

        ' Update Labels
        Label5.Text = "Total Income: R" & totalIncome.ToString("F2")
        Label6.Text = "Total Expenses: R" & totalExpenses.ToString("F2")
        Label8.Text = "Remaining: R" & remaining.ToString("F2")
        Label14.Text = "Used: " & usedPercent.ToString("F0") & "%"

        ' Budget alerts and panel behavior
        If totalExpenses >= 0.8 * budgetLimit AndAlso totalExpenses < budgetLimit Then
            Panel4.Visible = blinkState
            Panel4.BackColor = Color.Red
            ShowToast("Warning: 80% of your budget is used.")
            Label13.Text = "You have exceeded 80%"
            Label13.ForeColor = Color.Red
            'FlashAppWindow()
        ElseIf totalExpenses >= budgetLimit Then
            Panel4.Visible = blinkState
            Panel4.BackColor = Color.Red
            ShowToast("Alert: Budget exceeded!")
            Label13.Text = "Budget Exceeded!"
            Label13.ForeColor = Color.Red
            'FlashAppWindow()
        Else
            ' Below 80% usage - budget is under control
            Panel4.Visible = False
            Label13.Text = "Budget is under control"
            Label13.ForeColor = Color.Green
        End If


        blinkState = Not blinkState


    End Sub

    Private Function GetTotalExpenses() As Double
        Try

            Using conn As OleDbConnection = Getconnection()
                conn.Open()
                Dim cmd As New OleDbCommand("SELECT Amount FROM Expense", conn)
                Dim result = cmd.ExecuteScalar()
                Return If(IsDBNull(result), 0, Convert.ToDouble(result))
            End Using
        Catch ex As Exception
            Debug.WriteLine("Error fetching expenses: " & ex.Message)
            Return 0
        End Try
        conn.Close()

    End Function

    Private Function GetTotalIncome() As Double
        Try
            Using conn As OleDbConnection = Getconnection()
                conn.Open()
                Dim cmd As New OleDbCommand("SELECT Amount FROM Expense", conn)
                Dim result = cmd.ExecuteScalar()
                Return If(IsDBNull(result), 0, Convert.ToDouble(result))
            End Using
        Catch ex As Exception
            Debug.WriteLine("Error fetching income: " & ex.Message)
            Return 0
        End Try
        conn.Close()

    End Function

    Private Sub ShowToast(message As String)
        NotifyIcon1.BalloonTipTitle = "Budget Alert"
        NotifyIcon1.BalloonTipText = message
        NotifyIcon1.ShowBalloonTip(3000)
    End Sub

    ' Flash taskbar to simulate vibration
    <DllImport("user32.dll")>
    Private Shared Function FlashWindow(hWnd As IntPtr, bInvert As Boolean) As Boolean
    End Function

    Private Sub FlashAppWindow()
        FlashWindow(Me.Handle, True)
    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub LoadChart()


        Using conn As OleDbConnection = Getconnection()
            Dim cmd As New OleDbCommand("SELECT Tags, SUM(Amount) AS Total FROM Expense GROUP BY Tags", conn)
            Dim reader As OleDbDataReader

            Chart2.Series.Clear()
            Dim series As New Series("Expense")
            series.ChartType = SeriesChartType.Bar ' Set to Bar chart

            Try
                conn.Open()
                reader = cmd.ExecuteReader()
                While reader.Read()
                    If Not IsDBNull(reader("Tags")) AndAlso Not IsDBNull(reader("Total")) Then
                        series.Points.AddXY(reader("Tags").ToString(), Convert.ToDouble(reader("Total")))
                    End If
                End While
                Chart2.Series.Add(series)
            Catch ex As Exception
                MessageBox.Show("Error loading chart: " & ex.Message)
            Finally
                conn.Close()
            End Try
        End Using
    End Sub

    Private Sub LoadUpcomingMeals()
        Dim query As String = "Select MealName, StartDate, Description FROM MealPlans WHERE EndDate >= StartDate"

        ' Fetch data from the database

        Using conn As New OleDbConnection(connectionString)

            Using cmd As New OleDbCommand(query, conn)

                conn.Open()

                Using reader As OleDbDataReader = cmd.ExecuteReader()

                    ' Clear existing controls

                    FlowLayoutPanel1.Controls.Clear()

                    ' Loop through the data and create controls for each meal

                    While reader.Read()

                        Dim mealName As String = reader("MealName").ToString()

                        Dim startDate As DateTime = Convert.ToDateTime(reader("StartDate"))

                        Dim Description As String = reader("Description").ToString()

                        ' Create a new Label for each meal

                        Dim lblMeal As New Label()

                        lblMeal.Text = $"{mealName} on {startDate.ToShortDateString()} - {Description}"

                        lblMeal.AutoSize = True

                        lblMeal.Margin = New Padding(10)

                        lblMeal.BorderStyle = BorderStyle.FixedSingle

                        lblMeal.ForeColor = Color.Black

                        ' Add the label to the FlowLayoutPanel

                        FlowLayoutPanel1.Controls.Add(lblMeal)

                    End While

                End Using

            End Using

        End Using
        conn.Close()
    End Sub

    Dim expiredGroceries As New Queue(Of String)
    Dim backupGroceries As New List(Of String)
    Dim groceryTimer As New Timer()

    Private Sub LoadExpiringGroceries()
        expiredGroceries.Clear()
        backupGroceries.Clear()
        Label16.Text = ""

        Dim query As String = "SELECT [ItemName], [ExpiryDate] FROM [Inventory]"
        Try
            Using conn As New OleDbConnection(connectionString)
                Using cmd As New OleDbCommand(query, conn)
                    Dim adapter As New OleDbDataAdapter(cmd)
                    Dim dt As New DataTable()
                    adapter.Fill(dt)
                    conn.Open()
                    Dim today As Date = DateTime.Today

                    For Each row As DataRow In dt.Rows
                        Dim groceryName As String = row("ItemName").ToString()
                        Dim expiryDate As Date = Convert.ToDateTime(row("ExpiryDate"))

                        If expiryDate < today Then
                            Dim message As String = $"Expired Grocery: {groceryName} (Expired on {expiryDate:dd MMM yyyy})"
                            expiredGroceries.Enqueue(message)
                            backupGroceries.Add(message)
                        End If
                    Next

                    If expiredGroceries.Count > 0 Then
                        groceryTimer.Interval = 2000 ' 2 seconds
                        AddHandler groceryTimer.Tick, AddressOf ShowNextExpiredGrocery
                        groceryTimer.Start()
                    Else
                        Label16.Text = "No expired groceries."
                        Label16.ForeColor = Color.Black
                    End If
                End Using
            End Using
            conn.Close()

        Catch ex As Exception
            Debug.WriteLine("Error: " & ex.Message)
            MessageBox.Show("Failed to load expired groceries.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ShowNextExpiredGrocery(sender As Object, e As EventArgs)
        If expiredGroceries.Count = 0 AndAlso backupGroceries.Count > 0 Then
            ' Reset the queue from backup to loop
            For Each item In backupGroceries
                expiredGroceries.Enqueue(item)
            Next
        End If

        If expiredGroceries.Count > 0 Then
            Label16.Text = expiredGroceries.Dequeue()
            Label16.ForeColor = Color.Red
        End If
    End Sub
    Public Sub PopulateListboxFromTasks(ByRef ListBox As ListBox)
        Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)

        Try
            Debug.WriteLine("populate listbox from tasks successful")

            ' Open the connection
            conn.Open()

            ' SQL query to get Tittle and Priority from Tasks table
            Dim query As String = "SELECT Title, AssignedTo FROM Tasks"
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            ' Clear ListBox and add results
            ListBox.Items.Clear()
            While reader.Read()
                If Not IsDBNull(reader("Title")) AndAlso Not IsDBNull(reader("AssignedTo")) Then
                    ListBox.Items.Add($"{reader("Title")} - {reader("AssignedTo")}")
                End If
            End While

            reader.Close()
        Catch ex As Exception
            Debug.WriteLine("Failed to populate ListBox from tasks")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("Error: " & ex.Message)
        Finally
            'If conn.State = ConnectionState.Open Then conn.Close()
            conn.Close()
        End Try


    End Sub

    ' Function to show the pie chart for chore status
    Private Sub ShowChoreStatusPieChart()
        ' Create a dictionary to store the counts for each status
        Dim statusCounts As New Dictionary(Of String, Integer)

        ' Create a connection to the database
        Using conn As New OleDbConnection(connectionString)
            Try
                conn.Open()
                ' Query to get the count of chores by their status
                Dim query As String = "SELECT Status, COUNT(*) FROM Chores GROUP BY Status"
                Using cmd As New OleDbCommand(query, conn)
                    Using reader As OleDbDataReader = cmd.ExecuteReader()
                        ' Loop through each row of the result set
                        While reader.Read()
                            Dim status As String = reader.GetString(0) ' Get the status
                            Dim count As Integer = reader.GetInt32(1) ' Get the count
                            ' Add the status and count to the dictionary
                            If statusCounts.ContainsKey(status) Then
                                statusCounts(status) += count
                            Else
                                statusCounts.Add(status, count)
                            End If
                        End While
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message)
                conn.Close()
            End Try

        End Using

        ' Clear any previous data on the chart
        Chart1.Series.Clear()
        Chart1.Titles.Clear()

        ' Create a new series for the pie chart
        Dim series As New Series("Chore Status")
        series.ChartType = SeriesChartType.Pie
        series.IsValueShownAsLabel = True ' Show the count as a label

        ' Add the data points for each status and count
        For Each kvp In statusCounts
            series.Points.AddXY(kvp.Key, kvp.Value)
        Next

        ' Add the series to the chart
        Chart1.Series.Add(series)
        ' Add a title to the chart
        Chart1.Titles.Add("Chore Status Summary")
    End Sub


    ' Photo slideshow variables
    Private photoList As New List(Of String)()
    Private currentPhotoIndex As Integer = 0
    Private WithEvents photoTimer As New Timer()



    Private Sub LoadRecentPhotos()
        photoList.Clear()
        Dim query As String = "SELECT TOP 4 FilePath FROM Photos "

        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
            Using cmd As New OleDbCommand(query, conn)
                conn.Open()
                Using reader As OleDbDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim path As String = reader("FilePath").ToString()
                        If IO.File.Exists(path) Then
                            photoList.Add(path)
                        End If
                    End While
                End Using
            End Using
        End Using
    End Sub

    Private Sub DisplayPhoto()
        If photoList.Count = 0 Then Exit Sub

        ' Clear and display the current photo
        FlowLayoutPanel2.Controls.Clear()
        Dim pb As New PictureBox()
        pb.Image = Image.FromFile(photoList(currentPhotoIndex))
        pb.SizeMode = PictureBoxSizeMode.StretchImage
        pb.Size = If(FlowLayoutPanel2.Size.Width > 0, FlowLayoutPanel2.Size, New Size(200, 150))
        FlowLayoutPanel2.Controls.Add(pb)
    End Sub

    Private Sub SetupTimer()
        photoTimer.Interval = 2000 ' 2 seconds
        AddHandler photoTimer.Tick, AddressOf PhotoTimer_Tick
        photoTimer.Start()
    End Sub

    Private Sub PhotoTimer_Tick(sender As Object, e As EventArgs)
        If photoList.Count = 0 Then Exit Sub

        ' Move to next photo
        currentPhotoIndex = (currentPhotoIndex + 1) Mod photoList.Count
        DisplayPhoto()
    End Sub


    Dim scheduleAlerts As New Queue(Of String)
    Dim alertTimer As New Timer()

    'Dim scheduleAlerts As New Queue(Of String)
    Dim backupScheduleAlerts As New List(Of String)
    'Dim alertTimer As New Timer()

    Private Sub LoadFamilyScheduleAlerts()

        scheduleAlerts.Clear()
        backupScheduleAlerts.Clear()
        Label19.Text = ""

        Using conn As OleDbConnection = Getconnection()
            conn.Open()
            Dim query As String = "SELECT EventType, DateOfEvent, AssignedTo FROM FamilySchedule"
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            While reader.Read()
                Dim eventText As String = $"Family Schedule:{reader("DateOfEvent"):dd MMM} - {reader("EventType")} ({reader("AssignedTo")})"
                scheduleAlerts.Enqueue(eventText)
                backupScheduleAlerts.Add(eventText)
            End While

            reader.Close()
            conn.Close()

            If scheduleAlerts.Count > 0 Then
                alertTimer.Interval = 2000 ' 2 seconds
                AddHandler alertTimer.Tick, AddressOf ShowNextAlert
                alertTimer.Start()
            Else
                Label19.Text = "No upcoming family events."
            End If
        End Using
    End Sub

    Private Sub ShowNextAlert(sender As Object, e As EventArgs)
        If scheduleAlerts.Count = 0 AndAlso backupScheduleAlerts.Count > 0 Then
            For Each alert In backupScheduleAlerts
                scheduleAlerts.Enqueue(alert)
            Next
        End If

        If scheduleAlerts.Count > 0 Then
            Label19.Text = scheduleAlerts.Dequeue()
        End If
    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick

        'Timer3.Stop()
        RunSearchAndBlink()

    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click

        Notifications.ShowDialog()

    End Sub

    Private Sub UpdateNotificationCount()

        StyleBadge()

        Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)

        Dim lastViewedChores As DateTime = GetLastViewed(conn, "Chores")

        Dim lastViewedExpenses As DateTime = GetLastViewed(conn, "Expenses")

        Dim totalCount As Integer = 0

        Dim cmdChores As New OleDbCommand("SELECT COUNT(*) FROM Chores WHERE DueDate > ? AND DueDate > ?", conn)

        cmdChores.Parameters.AddWithValue("?", Date.Today)

        cmdChores.Parameters.AddWithValue("?", lastViewedChores)

        Dim cmdExpenses As New OleDbCommand("SELECT COUNT(*) FROM Expense WHERE Amount > ? AND DateOfexpenses > ?", conn)

        cmdExpenses.Parameters.AddWithValue("?", 1000)

        cmdExpenses.Parameters.AddWithValue("?", lastViewedExpenses)

        conn.Open()

        Dim choreCount As Integer = CInt(cmdChores.ExecuteScalar())

        Dim expenseCount As Integer = CInt(cmdExpenses.ExecuteScalar())

        conn.Close()

        totalCount = choreCount + expenseCount

        If totalCount > 0 Then

            lblbadge.Text = totalCount.ToString()

            lblbadge.Visible = True

        Else

            lblbadge.Visible = False

        End If

    End Sub

    Private Function GetLastViewed(con As OleDbConnection, viewType As String) As DateTime

        Dim cmd As New OleDbCommand("SELECT LastViewed FROM NotificationStatus WHERE ViewType = ?", conn)

        cmd.Parameters.AddWithValue("?", viewType)

        conn.Open()

        Dim result As Object = cmd.ExecuteScalar()

        conn.Close()

        If result IsNot Nothing Then

            Return CDate(result)

        Else

            Return DateTime.MinValue

        End If

    End Function

    Private Sub btnInAppMessages_Click(sender As Object, e As EventArgs) Handles btnInAppMessages.Click

        In_App_Message.ShowDialog()

        UpdateNotificationCount() ' Refresh count after closing

    End Sub

    Private Sub StyleBadge()

        lblbadge.Width = 20

        lblbadge.Height = 20

        Dim path As New Drawing2D.GraphicsPath()

        path.AddEllipse(0, 0, lblbadge.Width, lblbadge.Height)

        lblbadge.Region = New Region(path)

    End Sub
    ' Declare at the top of the form

    Dim currentButtonToBlink As Button = Nothing
    Dim formToOpen As Form = Nothing
    Dim blinkCounter As Integer = 0
    Dim delayCounter As Integer = 0
    Dim delayBeforeBlinking As Integer = 4 ' 4 ticks = ~1.2 sec (300ms x 4)
    Dim totalBlinkTicks As Integer = 10    ' ~3 sec of blinking
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        PrepareSearch()
    End Sub

    ' === Reset everything when user types ===
    Private Sub PrepareSearch()
        Timer3.Stop()
        delayCounter = 0
        blinkCounter = 0
        currentButtonToBlink = Nothing
        formToOpen = Nothing
        Timer3.Start()
    End Sub

    ' === Core subroutine for delay, blink, and open form ===
    Private Sub RunSearchAndBlink()
        ' Wait for 2 seconds
        If delayCounter < delayBeforeBlinking Then
            delayCounter += 1
            Return
        End If

        ' Set up blinking on first tick
        If blinkCounter = 0 Then
            Dim keyword As String = TextBox1.Text.ToLower()

            If keyword.Contains("mealplan") Then
                currentButtonToBlink = Button14
                formToOpen = New MealPlan()
            ElseIf keyword.Contains("expense") Then
                currentButtonToBlink = Button12
                formToOpen = New Expense()
            ElseIf keyword.Contains("inventory") Then
                currentButtonToBlink = Button15
                formToOpen = New Inventory()
            ElseIf keyword.Contains("chore") Then
                currentButtonToBlink = Button11
                formToOpen = New chores()
            ElseIf keyword.Contains("photo") Then
                currentButtonToBlink = Button8
                formToOpen = New PhotoGallery()
            ElseIf keyword.Contains("task") Then
                currentButtonToBlink = Button7
                formToOpen = New Task_Management
            ElseIf keyword.Contains("personnel") Then
                currentButtonToBlink = Button5
                formToOpen = New Personnel()
            ElseIf keyword.Contains("notification") Then
                currentButtonToBlink = Button16
                formToOpen = New Notifications()
            ElseIf keyword.Contains("grocery") Then
                currentButtonToBlink = Button13
                formToOpen = New Grocery_Items()
            ElseIf keyword.Contains("status") Then
                currentButtonToBlink = btnInAppMessages
                formToOpen = New In_App_Message()
            ElseIf keyword.Contains("family") Then
                currentButtonToBlink = Button6
                formToOpen = New Family_Schedule()
            ElseIf keyword.Contains("budget") Then
                currentButtonToBlink = Button17
                formToOpen = New Budget()
            Else
                Timer3.Stop()
                Exit Sub
            End If
        End If

        ' Toggle button color
        If currentButtonToBlink IsNot Nothing Then
            If blinkState Then
                currentButtonToBlink.BackColor = Color.LightGreen
            Else
                currentButtonToBlink.BackColor = Color.Red
            End If
            blinkState = Not blinkState
            blinkCounter += 1
        End If

        ' After blinking, reset and open
        If blinkCounter >= totalBlinkTicks Then
            Timer3.Stop()
            If currentButtonToBlink IsNot Nothing Then
                currentButtonToBlink.BackColor = SystemColors.Control
            End If
            If formToOpen IsNot Nothing Then
                MessageBox.Show("Opening " & formToOpen.Name, "Search Result", MessageBoxButtons.OK, MessageBoxIcon.Information)
                formToOpen.Show()
            End If
        End If

    End Sub

    Dim speedLevels As New List(Of String) From {
    "Fast", "Medium", "Slow", "Disconnected"
}
    Dim rand As New Random()

    Private Sub ShowInternetSpeed()
        Dim index As Integer = rand.Next(speedLevels.Count)
        Dim speed As String = speedLevels(index)
        Label23.Text = "Internet: " & speed

        Select Case speed
            Case "Fast"
                Label23.ForeColor = Color.Green
            Case "Medium"
                Label23.ForeColor = Color.Orange
            Case "Slow"
                Label23.ForeColor = Color.Red
            Case "Disconnected"
                Label23.ForeColor = Color.Gray
        End Select
    End Sub

    Private Sub ToolTip1_Popup(sender As Object, e As PopupEventArgs) Handles ToolTip1.Popup

    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick

    End Sub

    Private Sub Chart2_Click(sender As Object, e As EventArgs) Handles Chart2.Click

    End Sub

    Private Sub Label28_Click(sender As Object, e As EventArgs) Handles Label28.Click
        Login.Show()
        Me.Hide()
    End Sub


    Private Sub LoadMemberChoreTaskSummary()
        FlowLayoutPanel3.Controls.Clear()

        Dim query As String = "
        SELECT FirstName, LastName,
            (SELECT COUNT(*) FROM Chores WHERE AssignedTo = (FirstName & ' ' & LastName)) AS ChoreCount,
            (SELECT COUNT(*) FROM Tasks WHERE AssignedTo = (FirstName & ' ' & LastName)) AS TaskCount
        FROM PersonalDetails"

        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
            Using cmd As New OleDbCommand(query, conn)
                conn.Open()
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim fname As String = reader("FirstName").ToString()
                        Dim lname As String = reader("LastName").ToString()
                        Dim fullName As String = fname & " " & lname
                        Dim chores As Integer = Convert.ToInt32(reader("ChoreCount"))
                        Dim tasks As Integer = Convert.ToInt32(reader("TaskCount"))

                        Dim lbl As New Label()
                        lbl.Text = $"{fullName}
-{chores} chore
-{tasks} task"
                        lbl.AutoSize = True
                        lbl.Font = New Font("Segoe UI", 10, FontStyle.Regular)
                        lbl.Margin = New Padding(5)
                        FlowLayoutPanel3.Controls.Add(lbl)
                    End While
                End Using
            End Using
        End Using
    End Sub

    Private Sub LoadAssignmentSummary()
        Dim connString As String = HouseHoldManagment_Module.connectionString

        Dim totalQuery As String = "SELECT COUNT(*) FROM PersonalDetails"
        Dim assignedQuery As String = "
        SELECT COUNT(*) FROM (
            SELECT DISTINCT FirstName & ' ' & LastName AS FullName FROM PersonalDetails
            WHERE (FirstName & ' ' & LastName) IN (SELECT AssignedTo FROM Tasks)
               OR (FirstName & ' ' & LastName) IN (SELECT AssignedTo FROM Chores)
        ) AS AssignedPeople
    "

        Using conn As New OleDbConnection(connString)
            conn.Open()

            Dim cmdTotal As New OleDbCommand(totalQuery, conn)
            Dim totalMembers As Integer = Convert.ToInt32(cmdTotal.ExecuteScalar())

            Dim cmdAssigned As New OleDbCommand(assignedQuery, conn)
            Dim assignedMembers As Integer = Convert.ToInt32(cmdAssigned.ExecuteScalar())

            Dim unassignedMembers As Integer = totalMembers - assignedMembers

            Label29.Text = $"Total Members: {totalMembers}"
            Label30.Text = $"Assigned Members: {assignedMembers}"
            Label31.Text = $"Unassigned Members: {unassignedMembers}"
        End Using
    End Sub

    Private Sub LoadUpcomingDates()
        Dim connString As String = HouseHoldManagment_Module.connectionString
        Dim query As String = "
        SELECT 'Task' AS Status, Title, DueDate
        FROM Tasks 
        WHERE DueDate >= Date() AND DueDate <= DateAdd('d', 3, Date())
        UNION ALL
        SELECT 'Chore' AS Status, Title, DueDate
        FROM Chores 
        WHERE DueDate >= Date() AND DueDate <= DateAdd('d', 3, Date())
        ORDER BY DueDate ASC
    "

        Using conn As New OleDbConnection(connString)
            conn.Open()
            Dim cmd As New OleDbCommand(query, conn)
            Using reader = cmd.ExecuteReader()
                ListBox1.Items.Clear()
                While reader.Read()
                    Dim itemType As String = reader("Status").ToString()
                    Dim title As String = reader("Title").ToString()
                    Dim dueDate As Date = Convert.ToDateTime(reader("DueDate"))
                    'Dim assignedTo As String = reader("AssignedTo").ToString()

                    Dim displayText As String = $"{itemType}: {title} - Due {dueDate.ToShortDateString()})"
                    ListBox1.Items.Add(displayText)
                End While

                If ListBox1.Items.Count = 0 Then
                    ListBox1.Items.Add("No tasks or chores due in next 3 days.")
                End If
            End Using
        End Using
    End Sub
    Private Sub LoadUpcomingBirthdays()
        Dim connString As String = HouseHoldManagment_Module.connectionString
        Dim query As String = " SELECT FirstName, LastName, DateOfBirth FROM PersonalDetails WHERE DateOfBirth IS NOT NULL "
        Dim today As Date = Date.Today
        Dim endDate As Date = today.AddDays(7)  ' Changed from 30 days to 5 months

        Using conn As New OleDbConnection(connString)
            conn.Open()
            Dim cmd As New OleDbCommand(query, conn)
            Using reader = cmd.ExecuteReader()
                ListBox2.Items.Clear()
                While reader.Read()
                    Dim firstName As String = reader("FirstName").ToString()
                    Dim lastName As String = reader("LastName").ToString()
                    Dim birthDate As Date = Convert.ToDateTime(reader("DateOfBirth"))

                    ' Adjust birthday year to this year for comparison
                    Dim nextBirthday As Date = New Date(today.Year, birthDate.Month, birthDate.Day)

                    ' If birthday already passed this year, consider next year
                    If nextBirthday < today Then
                        nextBirthday = nextBirthday.AddYears(1)
                    End If

                    If nextBirthday >= today AndAlso nextBirthday <= endDate Then
                        Dim daysLeft As Integer = (nextBirthday - today).Days
                        ListBox2.Items.Add($"{firstName} {lastName} - Birthday in {daysLeft} day(s) on {nextBirthday.ToString("MMMM dd")}")
                    End If
                End While

                If ListBox2.Items.Count = 0 Then
                    ListBox2.Items.Add("No upcoming birthdays in the next 7 Days.")
                End If
            End Using
        End Using
    End Sub

    Private Sub CheckDatabaseConnection()
        'Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=your_database_path.accdb;"
        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
            Try
                conn.Open()
                Label36.Text = "Database connected successfully."
                Label36.ForeColor = Color.Green
            Catch ex As Exception
                Label36.Text = "Database Connection failed: " & ex.Message
                Label36.ForeColor = Color.Red
            Finally
                conn.Close()
            End Try
        End Using
    End Sub

    Private Sub LoadRecentGroceries()
        Try
            conn.Open()
            Dim query As String = "SELECT ItemName, Quantity, PurchaseDate FROM GroceryItems WHERE PurchaseDate >= Date() - 7 ORDER BY PurchaseDate DESC"
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            ListBox3.Items.Clear()

            If reader.HasRows Then
                While reader.Read()
                    Dim itemInfo As String = $"{reader("ItemName")} - Qty: {reader("Quantity")} - {CDate(reader("PurchaseDate")).ToShortDateString()}"
                    ListBox3.Items.Add(itemInfo)
                End While
            Else
                ListBox3.Items.Add("No groceries added in the last 7 days.")
            End If

        Catch ex As Exception
            MessageBox.Show("Error loading recent groceries: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub Label23_Click(sender As Object, e As EventArgs) Handles Label23.Click

    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class