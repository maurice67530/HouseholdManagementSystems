Imports System.IO
Imports System.Data.OleDb
Imports System.Net.Mail
Imports System.Net
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.DataVisualization.Charting

Public Class Dashboard
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Mudzunga\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"





    Private Sub Dashboard_Load(sender As Object, e As EventArgs) Handles MyBase.Load






        LoadChoresStatus()
        LoadChart()
        LoadUpcomingMeals()
        LoadExpiringGroceries()
        PopulateListboxFromTasks(Tasks)
        ShowChoreStatusPieChart()
        'LoadRecentPhotos()
        'DisplayPhoto()
        'SetupTimer()



        ' Show current date, month, and time
        Label17.Text = DateTime.Now.ToString("dddd, dd MMMM yyyy - hh:mm tt")
        'Label17.Location = New Point(10, 10) 'Adjust position as needed
        Label17.AutoSize = True
        Label17.Font = New Font("Arial", 10, FontStyle.Bold)
        Label17.ForeColor = Color.Blue

        Dim hour As Integer = DateTime.Now.Hour
        If hour < 12 Then
            Label18.Text = "Good morning!"
        ElseIf hour < 18 Then
            Label18.Text = "Good afternoon!"
        Else
            Label18.Text = "Good evening!"
        End If









        'DisplayPhoto()
        'SetupTimer()
        'LoadRecentPhotos()
        'SetupCharts()
        'LoadChoresStatus()

        'LoadUpcomingMeals()
        'UpdateBudgetStatus()
        'LoadChartData()

        ToolTip1.SetToolTip(Button7, "Task")
        ToolTip1.SetToolTip(Button15, "Inventory")
        ToolTip1.SetToolTip(Button12, "Expense")
        ToolTip1.SetToolTip(Button11, "Chores")
        ToolTip1.SetToolTip(Button14, "MealPlan")
        ToolTip1.SetToolTip(Button13, "GroceryItem")
        ToolTip1.SetToolTip(Button9, "Notification")
        ToolTip1.SetToolTip(Button5, "Personel")
        ToolTip1.SetToolTip(Button8, "PhotoGallery")
        ToolTip1.SetToolTip(Button6, "Family Event")

        Timer1.Interval = 100
        Timer1.Start()
        If photoList.Count > 0 Then
            DisplayPhoto() ' Show the first photo
            SetupTimer() ' Start timer for slideshow
        End If

        Timer2.Interval = 200
        Timer2.Start()



        LoadFamilyScheduleAlerts()


        'LoadExpensesData()
    End Sub

    ''Set up Budget Status And Chores Status charts
    'Private Sub SetupCharts()
    '    'Chores Status - Pie Chart
    '    Chart2.Series.Clear()
    '    Chart2.Series.Add("Chores")
    '    Chart2.Series("Chores").Points.AddXY("Completed", 0)
    '    Chart2.Series("Chores").Points.AddXY("In progress", 1)
    '    Chart2.Series("Chores").Points.AddXY("Not Started", 0)
    '    Chart2.Series("Chores").IsValueShownAsLabel = True
    '    ''Chart1.Series("Chores").ChartType = series1.Pie
    'End Sub













    'Private Sub LoadChartData()

    '    ' update this connection string based  on my database confirguration

    '    Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= C:\Users\Mudzunga\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"

    '    Dim query As String = "SELECT [Amount], [Frequency] FROM [Expense]"

    '    Using conn As New OleDbConnection(connectionString)

    '        Dim command As New OleDbCommand(query, conn)

    '        conn.Open()

    '        Using reader As OleDbDataReader = command.ExecuteReader

    '            While reader.Read

    '                ' assuming ColumnX is a string (category)  and columnY is numeric value

    '                Dim personnel As String = reader("Frequency").ToString

    '                Dim Budget As String = reader("Amount").ToString

    '                ' add points to the chart; chage the series name added

    '                Chart1.Series("Expense").Points.AddXY(personnel, Budget)

    '            End While

    '        End Using

    '    End Using

    '    Chart1.ChartAreas(0).AxisX.Title = "Frequency"

    '    Chart1.ChartAreas(0).AxisY.Title = "Amount"

    'End Sub

    'Public Sub PopulateListboxFromChores(ByRef Listbox As ListBox)

    '    Dim conn As New OleDbConnection(connectionString)

    '    Try

    '        Debug.WriteLine("populate listbox successful")

    '        'open the database connection

    '        conn.Open()

    '        'retrieve the firstname and surname columns from the personaldetails tabel

    '        Dim query As String = "SELECT ID, Status,Title FROM Chores"

    '        Dim cmd As New OleDbCommand(query, conn)

    '        Dim reader As OleDbDataReader = cmd.ExecuteReader()

    '        'bind the retrieved data to the combobox

    '        ListBox1.Items.Clear()

    '        While reader.Read()

    '            ListBox1.Items.Add($"{reader("ID")} {reader("Status")} {reader("Title")}")

    '        End While

    '        'close the database

    '        reader.Close()

    '    Catch ex As Exception

    '        'handle any exeptions that may occur  

    '        Debug.WriteLine("failed to populate ListBox")

    '        Debug.WriteLine($"Stack Trace: {ex.StackTrace}")

    '        MessageBox.Show($"Error: {ex.StackTrace}")

    '    Finally

    '        'close the database connection

    '        If conn.State = ConnectionState.Open Then

    '            conn.Close()

    '        End If

    '    End Try

    'End Sub



    'Private photoList As New List(Of String)() ' List to store photo paths

    'Private currentPhotoIndex As Integer = 0

    'Private WithEvents photoTimer As New Timer()

    'Private Sub LoadRecentPhotos()

    '    photoList.Clear()

    '    Dim query As String = "SELECT TOP 5 FilePath FROM Photos ORDER BY DateAdded "

    '    Using conn As New OleDbConnection(connectionString)

    '        Using cmd As New OleDbCommand(query, conn)

    '            conn.Open()

    '            Using reader As OleDbDataReader = cmd.ExecuteReader()

    '                While reader.Read()

    '                    photoList.Add(reader("FilePath").ToString())

    '                End While

    '            End Using

    '        End Using

    '    End Using

    'End Sub

    'Private Sub DisplayPhoto()

    '    'If photoList.Count > 0 Then

    '    '    FlowLayoutPanel2.Controls.Clear() ' Clear previous image

    '    '    Dim pb As New PictureBox()

    '    '    pb.Image = Image.FromFile(photoList(currentPhotoIndex))

    '    '    pb.SizeMode = PictureBoxSizeMode.StretchImage ' Set stretch mode

    '    '    pb.Size = FlowLayoutPanel2.Size ' Match panel size

    '    '    FlowLayoutPanel2.Controls.Add(pb) ' Add to FlowLayoutPanel

    '    'End If

    'End Sub

    'Private Sub SetupTimer()

    '    photoTimer.Interval = 2000 ' 2 seconds

    '    AddHandler photoTimer.Tick, AddressOf PhotoTimer_Tick

    '    photoTimer.Start()

    'End Sub

    'Private Sub PhotoTimer_Tick(sender As Object, e As EventArgs)

    '    If photoList.Count > 0 Then

    '        currentPhotoIndex = (currentPhotoIndex + 1) Mod photoList.Count ' Loop through photos

    '        DisplayPhoto()

    '    End If

    'End Sub
    'Private Sub LoadUpcomingMeals()
    '    Dim query As String = "Select MealName, StartDate, Description FROM MealPlans WHERE EndDate >= StartDate"

    '    ' Fetch data from the database

    '    Using conn As New OleDbConnection(connectionString)

    '        Using cmd As New OleDbCommand(query, conn)

    '            conn.Open()

    '            Using reader As OleDbDataReader = cmd.ExecuteReader()

    '                ' Clear existing controls

    '                FlowLayoutPanel1.Controls.Clear()

    '                ' Loop through the data and create controls for each meal

    '                While reader.Read()

    '                    Dim mealName As String = reader("MealName").ToString()

    '                    Dim startDate As DateTime = Convert.ToDateTime(reader("StartDate"))

    '                    Dim Description As String = reader("Description").ToString()

    '                    ' Create a new Label for each meal

    '                    Dim lblMeal As New Label()

    '                    lblMeal.Text = $"{mealName} on {startDate.ToShortDateString()} - {Description}"

    '                    lblMeal.AutoSize = True

    '                    lblMeal.Margin = New Padding(10)

    '                    lblMeal.BorderStyle = BorderStyle.FixedSingle

    '                    lblMeal.ForeColor = Color.Black

    '                    ' Add the label to the FlowLayoutPanel

    '                    FlowLayoutPanel1.Controls.Add(lblMeal)

    '                End While

    '            End Using

    '        End Using

    '    End Using

    'End Sub

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

        Label15.Text = $"   Chores: 
  -Completed: {completed}
  -In Progress:{inProgress}
  -Not Started:{notStarted}"
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

    Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles Button9.Click
        Notifications.ShowDialog()
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

        Using connection As New OleDbConnection(HouseHoldManagment_Module.connectionString)

            Try

                connection.Open()

                Using command As New OleDbCommand(query, connection)

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

        Using connection As New OleDbConnection(HouseHoldManagment_Module.connectionString)

            Try

                connection.Open()

                Using command As New OleDbCommand(query, connection)

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
            Using con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Mudzunga\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb")
                con.Open()
                Dim cmd As New OleDbCommand("SELECT SUM(Amount) FROM Expense", con)
                Dim result = cmd.ExecuteScalar()
                Return If(IsDBNull(result), 0, Convert.ToDouble(result))
            End Using
        Catch ex As Exception
            Debug.WriteLine("Error fetching expenses: " & ex.Message)
            Return 0
        End Try
    End Function

    Private Function GetTotalIncome() As Double
        Try
            Using con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Mudzunga\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb")
                con.Open()
                Dim cmd As New OleDbCommand("SELECT SUM(Totalincome) FROM Expense", con)
                Dim result = cmd.ExecuteScalar()
                Return If(IsDBNull(result), 0, Convert.ToDouble(result))
            End Using
        Catch ex As Exception
            Debug.WriteLine("Error fetching income: " & ex.Message)
            Return 0
        End Try
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

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click

    End Sub

    Private Sub LoadChart()


        Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Mudzunga\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb")
        Dim cmd As New OleDbCommand("SELECT Tags, SUM(Amount) AS Total FROM Expense GROUP BY Tags", con)
        Dim reader As OleDbDataReader

        Chart2.Series.Clear()
        Dim series As New Series("Expense")
        series.ChartType = SeriesChartType.Bar ' Set to Bar chart

        Try
            con.Open()
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
            con.Close()
        End Try

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

    End Sub

    Private Sub LoadExpiringGroceries()




        ' Define query to get expiry dates
        Dim query As String = "SELECT [ItemName], [ExpiryDate] FROM [Inventory]"

        Try
            ' Connect to database
            Using con As New OleDbConnection(connectionString)
                Using cmd As New OleDbCommand(query, con)
                    Dim adapter As New OleDbDataAdapter(cmd)
                    Dim dt As New DataTable()
                    adapter.Fill(dt)

                    ' Clear previous text
                    Label16.Text = ""

                    ' Check if data exists
                    Dim displayText As String = ""
                    Dim today As Date = DateTime.Today

                    ' Loop through each record and filter expired items
                    For Each row As DataRow In dt.Rows
                        Dim groceryName As String = row("ItemName").ToString()
                        Dim expiryDate As Date = Convert.ToDateTime(row("ExpiryDate"))


                        ' Only add groceries that have already expired (expiry date before today)
                        If expiryDate < today Then
                            displayText &= groceryName & " Expired on " & expiryDate.ToShortDateString() & Environment.NewLine
                        End If
                    Next

                    ' Display expired groceries
                    If displayText <> "" Then
                        Label16.Text = displayText
                        Label16.ForeColor = Color.Red ' Show expired items in red color
                    Else
                        Label16.Text = "No expired groceries."
                        Label16.ForeColor = Color.Black
                    End If
                End Using
            End Using

        Catch ex As Exception
            Debug.WriteLine("Error: " & ex.Message)
            MessageBox.Show("Failed to load expired groceries.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try



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
            If conn.State = ConnectionState.Open Then conn.Close()
        End Try
    End Sub

    ' Function to show the pie chart for chore status
    Private Sub ShowChoreStatusPieChart()
        ' Create a dictionary to store the counts for each status
        Dim statusCounts As New Dictionary(Of String, Integer)

        ' Create a connection to the database
        Using con As New OleDbConnection(connectionString)
            Try
                con.Open()
                ' Query to get the count of chores by their status
                Dim query As String = "SELECT Status, COUNT(*) FROM Chores GROUP BY Status"
                Using cmd As New OleDbCommand(query, con)
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
        photoTimer.Interval = 100 ' 2 seconds
        AddHandler photoTimer.Tick, AddressOf PhotoTimer_Tick
        photoTimer.Start()
    End Sub

    Private Sub PhotoTimer_Tick(sender As Object, e As EventArgs)
        If photoList.Count > 0 Then
            currentPhotoIndex = (currentPhotoIndex + 1) Mod photoList.Count ' Loop through photos
            DisplayPhoto()
        End If

    End Sub

    Private Sub LoadFamilyScheduleAlerts()
        Label19.Text = ""
        Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Mudzunga\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb")

        con.Open()
        Dim query As String = "SELECT EventType, DateOfEvent, AssignedTo FROM FamilySchedule"
        Dim cmd As New OleDbCommand(query, con)
        cmd.Parameters.AddWithValue("?", Date.Today)
        cmd.Parameters.AddWithValue("?", Date.Today.AddDays(5))

        Dim reader As OleDbDataReader = cmd.ExecuteReader()
        While reader.Read()
            Dim eventText As String = $" Family Schedule: {reader("DateOfEvent"):dd MMM} - {reader("EventType")} ({reader("AssignedTo")})"
            Label19.Text &= eventText & vbCrLf
        End While



        reader.Close()
        con.Close()
    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        blinkState = Not blinkState
        Label19.Visible = blinkState
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        In_App_Message.ShowDialog()
    End Sub
End Class