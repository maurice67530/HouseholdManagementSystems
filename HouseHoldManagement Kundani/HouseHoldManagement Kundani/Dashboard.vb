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



    Private Sub Dashboard_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblbadge.Region = New Region(New Drawing2D.GraphicsPath())
        Dim gp As New Drawing.Drawing2D.GraphicsPath()
        gp.AddEllipse(0, 0, lblbadge.Width, lblbadge.Height)
        lblbadge.Region = New Region(gp)
        StyleBadge()


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

    End Sub


    Private Sub LoadChoresStatus()

        Dim completed As Integer = 0, inProgress As Integer = 0, notStarted As Integer = 0
        Dim query As String = "SELECT Status, COUNT(*) FROM Chores GROUP BY Status"

        Using conn As New OleDbConnection(connectionString), cmd As New OleDbCommand(query, conn)
            conn.Open()
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Select Case reader("Status").ToString()
                        Case "Completed"
                            completed = Convert.ToInt32(reader(1))
                        Case "In progress"
                            inProgress = Convert.ToInt32(reader(1))
                        Case "Not Started"
                            notStarted = Convert.ToInt32(reader(1))
                    End Select

                End While

            End Using

        End Using
        conn.Close()
        Label15.Text = $"   Chores: -Completed: {completed} -In Progress:{inProgress} -Not Started:{notStarted}"
    End Sub
    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        Task_Management.ShowDialog()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Grocery_Items.ShowDialog()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Chores.ShowDialog()
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

    'Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles btnInAppMessages.Click
    '    Using con As OleDbConnection = Getconnection()


    '        Dim FullNames = TextBox2.Text.Trim()


    '        Dim cmd As New OleDbCommand("SELECT FullNames FROM Users WHERE Username = ? AND [Password] = ?", con)
    '        cmd.Parameters.AddWithValue("?", FullNames)


    '        con.Open()
    '        Dim reader As OleDbDataReader = cmd.ExecuteReader()

    '        If reader.Read() Then
    '            Dim family As String = reader("FullNames").ToString()


    '            MessageBox.Show("Login successful. Family: " & family, "Welcome!", MessageBoxButtons.OK, MessageBoxIcon.Information)


    '            In_App_Message.TextBox2.Text = FullNames


    '            In_App_Message.ShowDialog()
    '            Me.Hide()
    '        Else
    '            MessageBox.Show("cannot show Notification.")
    '        End If

    '        con.Close()
    '    End Using
    'End Sub

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

        Dim BudgetLimit As Decimal = 150169

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

                MessageBox.Show("Alert! you have used more that 80% of your budget", "Budget Alert", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            End If


            ' Send the email with the expired items

            Dim messageBody As String = $"Alert! Budget Alert:{vbCrLf}{vbCrLf}{BudgetLimit}"

            SendEmail("nethonondamudzunga45@gmail.com", "Budget Alert", messageBody)

            ' Notify that the email was sent

            MessageBox.Show("Budget Alert Alert Sent Successfully!", "Budget Alert", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception

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

            Dim smtpClient As New SmtpClient("smtp.gmail.com") With {.Port = 587, .EnableSsl = True, .Credentials = New NetworkCredential("nethonondamudzunga45@gmail.com", "slwo xavj lool amzu")}


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

    Dim budgetLimit As Double = 150300
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

        ' Budget alerts and panel blink
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
                Dim cmd As New OleDbCommand("SELECT SUM(Amount) FROM Expense", conn)
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
                Dim cmd As New OleDbCommand("SELECT SUM(Totalincome) FROM Expense", conn)
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

    Private photoList As New List(Of String)() ' List to store photo paths
    Private currentPhotoIndex As Integer = 0
    Private WithEvents photoTimer As New Timer()
    Private Sub LoadRecentPhotos()

        photoList.Clear()
        Dim query As String = "SELECT TOP 4 FilePath FROM Photos ORDER BY DateAdded"
        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
            Using cmd As New OleDbCommand(query, conn)
                conn.Open()
                Using reader As OleDbDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        photoList.Add(reader("FilePath").ToString())

                    End While
                End Using
            End Using
        End Using
        conn.Close()

    End Sub
    Private Sub DisplayPhoto()
        If photoList.Count > 0 Then
            FlowLayoutPanel2.Controls.Clear() ' Clear previous image
            Dim pb As New PictureBox()
            pb.Image = Image.FromFile(photoList(currentPhotoIndex))
            pb.SizeMode = PictureBoxSizeMode.StretchImage ' Set stretch mode
            pb.Size = FlowLayoutPanel2.Size ' Match panel size
            FlowLayoutPanel2.Controls.Add(pb) ' Add to FlowLayoutPanel
        End If
    End Sub

    Private Sub SetupTimer()
        photoTimer.Interval = 1000 ' 2 seconds
        AddHandler photoTimer.Tick, AddressOf PhotoTimer_Tick
        photoTimer.Start()
    End Sub

    Private Sub PhotoTimer_Tick(sender As Object, e As EventArgs)
        If photoList.Count > 0 Then
            currentPhotoIndex = (currentPhotoIndex + 1) Mod photoList.Count ' Loop through photos
            DisplayPhoto()
        End If

    End Sub
    Dim scheduleAlerts As New Queue(Of String)
    Dim alertTimer As New Timer()

    'Private Sub LoadFamilyScheduleAlerts()
    '    scheduleAlerts.Clear()
    '    Label19.Text = ""

    '    Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Mudzunga\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb")
    '    con.Open()
    '    Dim query As String = "SELECT EventType, DateOfEvent, AssignedTo FROM FamilySchedule"
    '    Dim cmd As New OleDbCommand(query, con)
    '    cmd.Parameters.AddWithValue("?", Date.Today)
    '    cmd.Parameters.AddWithValue("?", Date.Today.AddDays(5))

    '    Dim reader As OleDbDataReader = cmd.ExecuteReader()
    '    While reader.Read()
    '        Dim eventText As String = $"Family Schedule: {reader("DateOfEvent"):dd MMM} - {reader("EventType")} ({reader("AssignedTo")})"
    '        scheduleAlerts.Enqueue(eventText)
    '    End While
    '    reader.Close()
    '    con.Close()

    '    If scheduleAlerts.Count > 0 Then
    '        alertTimer.Interval = 2000 ' 2 seconds
    '        AddHandler alertTimer.Tick, AddressOf ShowNextAlert
    '        alertTimer.Start()
    '    Else
    '        Label19.Text = "No upcoming family events."
    '    End If
    'End Sub

    'Private Sub ShowNextAlert(sender As Object, e As EventArgs)
    '    If scheduleAlerts.Count > 0 Then
    '        Label19.Text = scheduleAlerts.Dequeue()
    '    Else
    '        alertTimer.Stop()
    '        RemoveHandler alertTimer.Tick, AddressOf ShowNextAlert
    '    End If
    'End Sub

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
                formToOpen = New Chores()
            ElseIf keyword.Contains("photo") Then
                currentButtonToBlink = Button8
                formToOpen = New PhotoGallery()
            ElseIf keyword.Contains("Task") Then
                currentButtonToBlink = Button7
                formToOpen = New Task_Management
            ElseIf keyword.Contains("Personnel") Then
                currentButtonToBlink = Button5
                formToOpen = New Personnel()
            ElseIf keyword.Contains("notification") Then
                currentButtonToBlink = Button16
                formToOpen = New Notifications()
            ElseIf keyword.Contains("grocery") Then
                currentButtonToBlink = Button13
                formToOpen = New Grocery_Items()
            ElseIf keyword.Contains("notificationstatus") Then
                currentButtonToBlink = btnInAppMessages
                formToOpen = New In_App_Message()
            ElseIf keyword.Contains("Family") Then
                currentButtonToBlink = Button6
                formToOpen = New Family_Schedule()
                'ElseIf keyword.Contains("budget") Then
                '    currentButtonToBlink = Button17
                '    formToOpen = New Budget()
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
End Class