Imports System.IO
Imports System.Data.OleDb
Imports System.Net.Mail
Imports System.Net
Public Class Dashboard
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Rinae\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"
    Private Sub Dashboard_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        DisplayPhoto()
        SetupTimer()
        LoadRecentPhotos()
        SetupCharts()
        LoadChoresStatus()

        LoadUpcomingMeals()
        UpdateBudgetStatus()
        LoadChartData()

        ToolTip1.SetToolTip(Button7, "Task")
        ToolTip1.SetToolTip(Button15, "Inventory")
        ToolTip1.SetToolTip(Button12, "Expense")
        ToolTip1.SetToolTip(Button11, "Chores")
        ToolTip1.SetToolTip(Button14, "MealPlan")
        ToolTip1.SetToolTip(Button6, "GroceryItem")
        ToolTip1.SetToolTip(Button9, "Notification")
        ToolTip1.SetToolTip(Button5, "Personel")
        ToolTip1.SetToolTip(Button8, "PhotoGallery")
        ToolTip1.SetToolTip(Button6, "Family Event")

        Timer1.Start()
        Timer1.Interval = 3000

    End Sub

    'Set up Budget Status And Chores Status charts
    Private Sub SetupCharts()
        'Chores Status - Pie Chart
        Chart2.Series.Clear()
        Chart2.Series.Add("Chores")
        Chart2.Series("Chores").Points.AddXY("Completed", 0)
        Chart2.Series("Chores").Points.AddXY("In progress", 1)
        Chart2.Series("Chores").Points.AddXY("Not Started", 0)
        Chart2.Series("Chores").IsValueShownAsLabel = True
        ''Chart1.Series("Chores").ChartType = series1.Pie
    End Sub

    Private Sub UpdateBudgetStatus()

        Dim query As String = "SELECT SUM(Amount) FROM Expense"

        Using conn As New OleDbConnection(connectionString)

            conn.Open()

            Dim cmd As New OleDbCommand(query, conn)

            Dim totalExpenses As Decimal = Convert.ToDecimal(cmd.ExecuteScalar())

            ' Assume you have a Label for Budget

            Label2.Text = "Total Expenses: R" & totalExpenses.ToString()

            ' Assuming a fixed budget, for example $500

            Dim budget As Decimal = 500563

            Label3.Text = "Budget Used: " & ((totalExpenses / budget) * 100).ToString("F2") & "%"

            ' Update a progress bar if you have one

            ProgressBar1.Value = CInt((totalExpenses / budget) * 100)

        End Using

    End Sub

    Private Sub LoadChartData()

        ' update this connection string based  on my database confirguration

        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= C:\Users\Rinae\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"

        Dim query As String = "SELECT [Amount], [Frequency] FROM [Expense]"

        Using conn As New OleDbConnection(connectionString)

            Dim command As New OleDbCommand(query, conn)

            conn.Open()

            Using reader As OleDbDataReader = command.ExecuteReader

                While reader.Read

                    ' assuming ColumnX is a string (category)  and columnY is numeric value

                    Dim personnel As String = reader("Frequency").ToString

                    Dim Budget As String = reader("Amount").ToString

                    ' add points to the chart; chage the series name added

                    Chart1.Series("Expense").Points.AddXY(personnel, Budget)

                End While

            End Using

        End Using

        Chart1.ChartAreas(0).AxisX.Title = "Frequency"

        Chart1.ChartAreas(0).AxisY.Title = "Amount"

    End Sub

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



    Private photoList As New List(Of String)() ' List to store photo paths

    Private currentPhotoIndex As Integer = 0

    Private WithEvents photoTimer As New Timer()

    Private Sub LoadRecentPhotos()

        photoList.Clear()

        Dim query As String = "SELECT TOP 5 FilePath FROM Photos ORDER BY DateAdded "

        Using conn As New OleDbConnection(connectionString)

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

        photoTimer.Interval = 2000 ' 2 seconds

        AddHandler photoTimer.Tick, AddressOf PhotoTimer_Tick

        photoTimer.Start()

    End Sub

    Private Sub PhotoTimer_Tick(sender As Object, e As EventArgs)

        If photoList.Count > 0 Then

            currentPhotoIndex = (currentPhotoIndex + 1) Mod photoList.Count ' Loop through photos

            DisplayPhoto()

        End If

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

        Label1.Text = $"   Chores: 
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






End Class