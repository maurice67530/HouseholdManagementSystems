Imports System.IO
Imports System.Data.OleDb
Public Class Dash
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Rinae\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"
    Private Sub Dash_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DisplayPhoto()
        SetupTimer()

        SetupCharts()
        LoadChoresStatus()

        LoadUpcomingMeals()
        UpdateBudgetStatus()
        LoadChartData()

        ToolTip1.SetToolTip(Button7, "Inventory")
        ToolTip1.SetToolTip(Button3, "Task")
        ToolTip1.SetToolTip(Button5, "Expense")
        ToolTip1.SetToolTip(Button2, "Chores")
        ToolTip1.SetToolTip(Button1, "MealPlan")
        ToolTip1.SetToolTip(Button6, "GroceryItem")
        ToolTip1.SetToolTip(Button4, "Person")
        ToolTip1.SetToolTip(Button9, "Photos")

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


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MealPlan.ShowDialog()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'Groceryy.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Chores.ShowDialog()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Inventory.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'DailyTask.ShowDialog()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Expense.ShowDialog()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Personnel.ShowDialog()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        PhotoGallery.ShowDialog()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Notifications.ShowDialog()
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
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Rinae\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"
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

                    Chart1.Series("Series1").Points.AddXY(personnel, Budget)

                End While

            End Using

        End Using

        Chart1.ChartAreas(0).AxisX.Title = "Frequency"

        Chart1.ChartAreas(0).AxisY.Title = "Amount"
    End Sub

    Public Sub PopulateListboxFromChores(ByRef Listbox As ListBox)

        Dim conn As New OleDbConnection(connectionString)
        Try
            Debug.WriteLine("populate listbox successful")
            'open the database connection
            conn.Open()

            'retrieve the firstname and surname columns from the personaldetails tabel
            Dim query As String = "SELECT ID, Status,Title FROM Chores"
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            'bind the retrieved data to the combobox
            ListBox1.Items.Clear()
            While reader.Read()
                ListBox1.Items.Add($"{reader("ID")} {reader("Status")} {reader("Title")}")
            End While

            'close the database
            reader.Close()
        Catch ex As Exception
            'handle any exeptions that may occur  
            Debug.WriteLine("failed to populate ListBox")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error: {ex.StackTrace}")

        Finally

            'close the database connection
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If

        End Try

    End Sub



    Private photoList As New List(Of String)() ' List to store photo paths
    Private currentPhotoIndex As Integer = 0
    Private WithEvents photoTimer As New Timer()
    Private Sub LoadRecentPhotos()

        photoList.Clear()
        Dim query As String = "SELECT TOP 5 FilePath FROM Photos ORDER BY DateTaken "
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
            pb.Size = FlowLayoutPanel1.Size ' Match panel size
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
End Class