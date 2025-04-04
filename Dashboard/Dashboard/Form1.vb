Imports System.Data.OleDb
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadChartData()
        SetupCharts()
    End Sub
    'Set up Budget Status And Chores Status charts
    Private Sub SetupCharts()
        'Chores Status - Pie Chart
        Chart2.Series.Clear()
        Chart2.Series.Add("Chores")
        Chart2.Series("Chores").Points.AddXY("Completed", 2)
        Chart2.Series("Chores").Points.AddXY("In progress", 1)
        Chart2.Series("Chores").Points.AddXY("Not Started", 3)
        Chart2.Series("Chores").IsValueShownAsLabel = True
        '    Chart2.Series("Chores").ChartType = Series1.Pie
    End Sub


    Private Sub LoadChartData()
        ' Update this connection string based  on my database confirguration
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Mulanga\Downloads\Household Management\New Microsoft Access Database.accdb"
        Dim query As String = "SELECT [Amount], [Person] FROM [Expenses]"

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

                    ' Chart1.Series("Series1").Points.AddXY(Person, Budget)

                End While
            End Using
        End Using
        'Chart1.ChartAreas(0).AxisX.Title = "Person"
        'Chart1.ChartAreas(0).AxisY.Title = "Amount"

    End Sub


    '   Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
    '   Dashboard.show()
    '    End Sub

    '    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '        Inventory.show()
    '    End Sub

    '    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    '        Task.show()
    '    End Sub

    '    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
    '        Expense.show()
    '    End Sub

    '    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
    '        Chores.show()
    '    End Sub
End Class
