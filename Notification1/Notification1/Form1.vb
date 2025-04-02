Imports System.IO
Imports System.Data.OleDb
Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Conn As New OleDbConnection(NotificationModule.connectionString)

        If DataGridView1.SelectedRows.Count = 0 Then
            MsgBox("Please select a notification to mark as read.", MsgBoxStyle.Exclamation, "No Selection")
            Return
        End If

        Conn.Open()

        For Each row As DataGridViewRow In DataGridView1.SelectedRows
            Dim notificationID As Integer = Convert.ToInt32(row.Cells("ID").Value)
            Dim query As String = "UPDATE Notifications SET IsRead = True WHERE ID = @ID"

            Using cmd As New OleDbCommand(query, Conn)
                cmd.Parameters.AddWithValue("@ID", notificationID)
                cmd.ExecuteNonQuery()
            End Using

            row.DefaultCellStyle.ForeColor = Color.Black ' Change UI for read status
        Next

        Conn.Close()
        MsgBox("Selected notifications marked as read!", MsgBoxStyle.Information, "Updated")

        ' Refresh unread count after marking as read
        CountUnreadNotifications()



    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Conn As New OleDbConnection(NotificationModule.connectionString)


        If DataGridView1.SelectedRows.Count = 0 Then
            MsgBox("Please select a notification to delete.", MsgBoxStyle.Exclamation, "No Selection")
            Return
        End If

        Conn.Open()

        For Each row As DataGridViewRow In DataGridView1.SelectedRows
            Dim notificationID As Integer = Convert.ToInt32(row.Cells("ID").Value)
            Dim query As String = "DELETE FROM Notifications WHERE ID = @ID"

            Using cmd As New OleDbCommand(query, Conn)
                cmd.Parameters.AddWithValue("@ID", notificationID)
                cmd.ExecuteNonQuery()
            End Using

            DataGridView1.Rows.Remove(row)
        Next

        Conn.Close()
        MsgBox("Selected notifications cleared!", MsgBoxStyle.Information, "Deleted")

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadNotifications()
        CountUnreadNotifications()
        CheckInventory()
        CheckOverdueTasks()
        CheckOverdueChores()
    End Sub
    Sub LoadNotifications()
        Dim Conn As New OleDbConnection(NotificationModule.connectionString)

        ListBox1.Items.Clear()

        Try
            ' Open Connection
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If

            Dim query As String = "SELECT ID, Message FROM Notifications ORDER BY DateCreated DESC"
            Dim cmd As New OleDbCommand(query, Conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            While reader.Read()
                Dim ID As Integer = reader("ID")
                Dim message As String = reader("Message")
                ListBox1.Items.Add(ID & " - " & message) ' Display ID & Message
            End While

            reader.Close()
            CountUnreadNotifications()

        Catch ex As Exception
            MessageBox.Show("Error loading notifications: " & ex.Message)

        Finally
            ' Close Connection
            If Conn.State = ConnectionState.Open Then
                Conn.Close()
            End If
        End Try
    End Sub
    Private Sub CountUnreadNotifications()
        Dim conn As New OleDbConnection(NotificationModule.connectionString)
        Try
            Dim query As String = "SELECT COUNT(*) FROM Notifications WHERE IsRead = False"
            Dim cmd As New OleDbCommand(query, conn)
            conn.Open()
            Dim unreadCount As Integer = Convert.ToInt32(cmd.ExecuteScalar())
            conn.Close()

            ' Update the label with the unread count
            Label2.Text = "Unread Notifications: " & unreadCount
        Catch ex As Exception
            Debug.WriteLine("Error counting unread notifications: " & ex.Message)
        End Try

    End Sub
    Dim blink As Boolean = False
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If blink Then
            Label3.Visible = True
        Else
            Label3.Visible = False
        End If
        blink = Not blink
    End Sub
    Public Sub CheckExpenseNotification()
        Dim totalExpenses As Double = 0
        Dim budgetLimit As Double = 6000
        Dim warningThreshold As Double = budgetLimit * 0.8 ' 80% of the budget

        Try
            ' Connect to the database
            Dim conn As New OleDbConnection(NotificationModule.connectionString)
            Dim cmd As New OleDb.OleDbCommand("SELECT SUM(Amount) FROM Expenses", conn)

            conn.Open()
            Dim result = cmd.ExecuteScalar()
            If Not IsDBNull(result) Then
                totalExpenses = Convert.ToDouble(result)
            End If
            conn.Close()

            ' Display warning if expenses exceed 80%
            If totalExpenses >= warningThreshold Then
                Label3.Text = "Warning! High Expenses!"
                Label3.ForeColor = Color.Red
                Timer1.Start() ' Start blinking effect
                ProgressBar1.Value = 100 ' Set progress to full
                ProgressBar1.ForeColor = Color.Red
            Else
                Label3.Text = "Budget Status: Normal"
                Label3.ForeColor = Color.Green
                Timer1.Stop()
                ProgressBar1.Value = (totalExpenses / budgetLimit) * 100
                ProgressBar1.ForeColor = Color.Green
            End If

        Catch ex As Exception
            Debug.WriteLine("Error: " & ex.Message)
        End Try
    End Sub
    ' Function to check for overdue tasks and send a notification
    Private Sub CheckOverdueTasks()
        Dim conn As New OleDbConnection(NotificationModule.connectionString)
        'Dim query As String = "SELECT TaskName, DueDate, AssignedTo FROM Tasks WHERE DueDate < @CurrentDate AND Status <> 'Completed'"

        'Using connection As New OleDbConnection(connectionString)
        '    Dim command As New OleDbCommand(query, connection)
        '    command.Parameters.AddWithValue("@CurrentDate", DateTime.Now) ' Current Date

        '    connection.Open()

        '    Dim reader As OleDbDataReader = command.ExecuteReader()

        '    ' Check if there are any overdue tasks
        '    If reader.HasRows Then
        '        While reader.Read()
        '            Dim taskName As String = reader("TaskName").ToString()
        '            Dim dueDate As DateTime = Convert.ToDateTime(reader("DueDate"))
        '            Dim assignedTo As Integer = Convert.ToInt32(reader("AssignedTo"))

        '            ' Trigger overdue notification
        '            Dim message As String = $"Overdue Task: {taskName} was due on {dueDate.ToShortDateString()}."
        '            SendNotification(message, "Task", assignedTo)
        '        End While
        '    End If

        '    reader.Close()
        'End Using

        Try
            Dim query As String = "SELECT TaskName FROM Tasks WHERE DueDate < Date() AND Status <> 'Completed'"
            Dim adapter As New OleDbDataAdapter(query, conn)
            Dim dt As New DataTable()
            adapter.Fill(dt)

            For Each row As DataRow In dt.Rows
                AddNotification($"Task '{row("TaskName")}' is overdue!", "Task")
            Next
        Catch ex As Exception
            Debug.WriteLine("Error checking overdue tasks: " & ex.Message)
        End Try
    End Sub


    ' Function to add a notification to the database
    Private Sub SendNotification(message As String, category As String, userId As Integer)
        Dim conn As New OleDbConnection(NotificationModule.connectionString)
        Dim query As String = "INSERT INTO Notifications (Message, DateCreated, Category, UserId, IsRead) VALUES (@Message, @DateCreated, @Category, @UserId, IsRead)"

        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(query, connection)
            command.Parameters.AddWithValue("@Message", message)
            command.Parameters.AddWithValue("@DateCreated", DateTime.Now)
            command.Parameters.AddWithValue("@Category", category)
            command.Parameters.AddWithValue("@UserId", userId)
            command.Parameters.AddWithValue("@IsRead", False)
            connection.Open()
            command.ExecuteNonQuery()
        End Using
    End Sub

    ' Function to check for overdue chores and send a notification
    Private Sub CheckOverdueChores()
        Dim conn As New OleDbConnection(NotificationModule.connectionString)
        Dim query As String = "SELECT ChoreName, DueDate, AssignedTo FROM Chores WHERE DueDate < @CurrentDate AND Status <> 'Completed'"

        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(query, connection)
            command.Parameters.AddWithValue("@CurrentDate", DateTime.Now) ' Current Date

            connection.Open()

            Dim reader As OleDbDataReader = command.ExecuteReader()

            ' Check if there are any overdue chores
            If reader.HasRows Then
                While reader.Read()
                    Dim choreName As String = reader("ChoreName").ToString()
                    Dim dueDate As DateTime = Convert.ToDateTime(reader("DueDate"))
                    Dim assignedTo As Integer = Convert.ToInt32(reader("AssignedTo"))

                    ' Trigger overdue notification
                    Dim message As String = $"Overdue Chore: {choreName} was due on {dueDate.ToShortDateString()}."
                    SendNotification(message, "Chore", assignedTo)
                End While
            End If

            reader.Close()
        End Using
    End Sub
    ' Function to check inventory quantity and send notification if below 5

    Public Sub CheckInventory()
        Dim conn As New OleDbConnection(NotificationModule.connectionString)
        Try
            Dim query As String = "SELECT ItemName, Quantity FROM Inventory WHERE Quantity < 5"
            Dim adapter As New OleDbDataAdapter(query, conn)
            Dim dt As New DataTable()
            adapter.Fill(dt)

            For Each row As DataRow In dt.Rows
                AddNotification($"Low stock alert: '{row("ItemName")}' is running low!", "Inventory")
            Next
        Catch ex As Exception
            Debug.WriteLine("Error checking inventory: " & ex.Message)
        End Try

    End Sub
    Public Sub AddNotification(message As String, category As String)
        Dim conn As New OleDbConnection(NotificationModule.connectionString)
        Try
            Dim query As String = "INSERT INTO Notifications (Message, Category, DateCreated, IsRead) VALUES (@Message, @Category, Now(), False)"
            Using cmd As New OleDbCommand(query, conn)
                cmd.Parameters.AddWithValue("@Message", message)
                cmd.Parameters.AddWithValue("@Category", category)
                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Debug.WriteLine("Error adding notification: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub
End Class
