Imports System.Data.OleDb
Public Class Dashboard
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

        LoadChoresStatus()

        LoadExpensesData()

        UpdateBudgetStatus()

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

    Private Sub LoadChoresStatus()

        Dim completed As Integer = 0, inProgress As Integer = 0, notStarted As Integer = 0
        Dim query As String = "SELECT Status, COUNT(*) FROM Chore GROUP BY Status"

        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString), cmd As New OleDbCommand(query, conn)
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

    Private Sub LoadExpensesData()

        Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
        Dim cmd As New OleDbCommand("SELECT Description, Amount As TotalExpense FROM Expense", conn)
        Dim dt As New DataTable()

        Try
            conn.Open()
            Dim adapter As New OleDbDataAdapter(cmd)
            adapter.Fill(dt) ' Load data into the DataTable

            ' Clear previous data
            Label2.Text = "Total Expenses: 0"
            Chart1.Series("Expenses").Points.Clear()
            Dim totalExpenses As Decimal = 0
            Dim uniqueTags As New Dictionary(Of String, Decimal) ' Load data into chart with clean tags
            For Each row As DataRow In dt.Rows

                Dim tag As String = row("Description").ToString().Trim()
                Dim totalExpense As Decimal = Convert.ToDecimal(row("TotalExpense"))
                totalExpenses += totalExpense


                ' Combine expenses for the same tag
                If uniqueTags.ContainsKey(tag) Then
                    uniqueTags(tag) += totalExpense
                Else
                    uniqueTags(tag) = totalExpense
                End If
            Next

            ' Check user role for displaying information
            ' If TextBox2.Text = "Finance" Or TextBox2.Text = "Admin" Then

            ' Update chart
            If Chart1.Series.IndexOf("Expense") <> -1 Then
                For Each kvp As KeyValuePair(Of String, Decimal) In uniqueTags
                    Chart1.Series("Expense").Points.AddXY(kvp.Key, kvp.Value)
                Next

            End If

            'End If
        Catch ex As Exception

            MessageBox.Show("Error loading expenses data: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

            conn.Close()

        End Try
    End Sub

    Private Sub UpdateBudgetStatus()

        Dim query As String = "SELECT SUM(Amount) FROM Expense"

        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
            conn.Open()
            Dim cmd As New OleDbCommand(query, conn)
            Dim totalExpenses As Decimal = Convert.ToDecimal(cmd.ExecuteScalar())
            ' Assume you have a Label for Budget
            Label2.Text = "Total Expenses: R" & totalExpenses.ToString()
            ' Assuming a fixed budget, for example $500
            Dim budget As Decimal = 24147
            Label3.Text = "Budget Used: " & ((totalExpenses / budget) * 100).ToString("F2") & "%"
            ' Update a progress bar if you have one
            ProgressBar1.Value = CInt((totalExpenses / budget) * 100)
        End Using
    End Sub

    Private Sub LoadChartData()
        ' Update this connection string based  on my database confirguration
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= C:\Users\Mulanga\Source\Repos\HouseholdManagementSystems\HMS.accdb"
        Dim query As String = "SELECT [Amount], [Person] FROM [Expense]"

        Using conn As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(query, conn)
            conn.Open()

            Using reader As OleDbDataReader = command.ExecuteReader
                While reader.Read
                    ' assuming ColumnX is a string (category)  and columnY is numeric value
                    'Chart1.Series.Add("BudgetStatus")
                    Dim Person As String = reader("Person").ToString
                    Dim Budget As String = reader("Amount").ToString

                    ' add points to the chart; chage the series name added

                    Chart1.Series("Series1").Points.AddXY(Person, Budget)

                End While
            End Using
        End Using
        Chart1.ChartAreas(0).AxisX.Title = "Person"
        Chart1.ChartAreas(0).AxisY.Title = "Amount"
    End Sub
End Class