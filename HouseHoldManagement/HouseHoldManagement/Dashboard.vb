﻿
Imports System.IO
Imports System.Data.OleDb
Public Class Dashboard
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= C:\Users\Rinae\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb "

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Inventory.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Task .ShowDialog()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Expense.ShowDialog()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Chores.ShowDialog()
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs)
        Login.ShowDialog()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        'Grocery.ShowDialog()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Personnel.ShowDialog()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        ' photo.ShowDialog()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        MealPlan.ShowDialog()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Notifications.ShowDialog()
    End Sub
    Private Sub Dashboard_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadUpcomingMeals()
        LoadChoresStatus()
        '  PopulateListboxFromChores(ListBox1)
        'LoadExpensesData()
        LoadChartData()
        SetupCharts()
        UpdateBudgetStatus()
        'LoadRecentPhotos()
        ' ToolTip1.SetToolTip(Button1, "Login")
        ToolTip1.SetToolTip(Button2, "Inventory")
        ToolTip1.SetToolTip(Button3, "Task")
        ToolTip1.SetToolTip(Button4, "Expense")
        ToolTip1.SetToolTip(Button5, "Chores")
        ToolTip1.SetToolTip(Button6, "MealPlan")
        ToolTip1.SetToolTip(Button7, "GroceryItem")
        ToolTip1.SetToolTip(Button8, "Person")
        ToolTip1.SetToolTip(Button9, "Photos")
        ToolTip1.SetToolTip(Button10, "Notification")

    End Sub
    'Set up Budget Status And Chores Status charts
    Private Sub SetupCharts()
        ' Chores Status - Pie Chart
        Chart2.Series.Clear()
        Chart2.Series.Add("Chores")
        Chart2.Series("Chores").Points.AddXY("Completed", 0)
        Chart2.Series("Chores").Points.AddXY("In progress", 1)
        Chart2.Series("Chores").Points.AddXY("Not Started", 0)
        Chart2.Series("Chores").IsValueShownAsLabel = True
        ''Chart1.Series("Chores").ChartType = series1.Pie
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
        Label2.Text = $"   Chores: 
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


    Dim increment As Integer = 0
    Private Sub UpdateBudgetStatus()

        Dim query As String = "SELECT SUM(Amount) FROM Expense"

        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Dim cmd As New OleDbCommand(query, conn)
            Dim totalExpenses As Decimal = Convert.ToDecimal(cmd.ExecuteScalar())
            ' Assume you have a Label for Budget
            Label5.Text = "Total Expenses: R" & totalExpenses.ToString()
            ' Assuming a fixed budget, for example $500
            Dim budget As Decimal = 500563
            Label6.Text = "Budget Used: " & ((totalExpenses / budget) * 100).ToString("F2") & "%"
            ' Update a progress bar if you have one
            ProgressBar2.Value = CInt((totalExpenses / budget) * 100)
        End Using
    End Sub

    Private Sub LoadChartData()
        ' update this connection string based  on my database confirguration
        Dim connectionString As String = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source= C:\Users\Rinae\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"
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

                    Chart1.Series("Expenses").Points.AddXY(personnel, Budget)

                End While
            End Using
        End Using
        Chart1.ChartAreas(0).AxisX.Title = "Frequency"
        Chart1.ChartAreas(0).AxisY.Title = "Amount"
    End Sub
    'Private photoList As New List(Of String)() ' List to store photo paths
    'Private currentPhotoIndex As Integer = 0
    'Private WithEvents photoTimer As New Timer()
    'Private Sub LoadRecentPhotos()

    '    photoList.Clear()
    '    Dim query As String = "SELECT TOP 5 FilePath FROM photo ORDER BY DateAdded "
    '    Using conn As New OleDbConnection(connectionString)
    '        Using cmd As New OleDbCommand(query, conn)
    '            conn.Open()
    '            Using reader As OleDbDataReader = cmd.ExecuteReader()
    '                While reader.Read()
    '                    photoList.Add(reader("filepath").ToString())
    '                End While
    '            End Using
    '        End Using
    '    End Using
    'End Sub
    'Private Sub DisplayPhoto()
    '    If photoList.Count > 0 Then
    '        FlowLayoutPanel1.Controls.Clear() ' Clear previous image
    '        Dim pb As New PictureBox()
    '        pb.Image = Image.FromFile(photoList(currentPhotoIndex))
    '        pb.SizeMode = PictureBoxSizeMode.StretchImage ' Set stretch mode
    '        pb.Size = FlowLayoutPanel1.Size ' Match panel size
    '        FlowLayoutPanel1.Controls.Add(pb) ' Add to FlowLayoutPanel
    '    End If
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
End Class