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
        'LoadNotifications()
        'CountUnreadNotifications()
        'CheckInventory()
        'CheckOverdueTasks()
        'CheckOverdueChores()

        Timer1.Start()
    End Sub
    Private Sub LoadNotifications()
        Dim conn As New OleDbConnection(NotificationModule.connectionString)
        Try
            Dim query As String = "SELECT ID, Message, Category, DateCreated, IsRead FROM Notifications ORDER BY DateCreated DESC"
            Dim adapter As New OleDbDataAdapter(query, conn)
            Dim dt As New DataTable()
            adapter.Fill(dt)
            DataGridView1.DataSource = dt

            ' Highlight unread notifications
            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not Convert.ToBoolean(row.Cells("IsRead").Value) Then
                    row.DefaultCellStyle.ForeColor = Color.Red
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine("Error loading notifications: " & ex.Message)
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

        CheckExpenses()
        CheckOverdueChores()
        CheckOverdueTasks()
        CheckInventory()
        LoadNotifications()
    End Sub
    Public Sub CheckExpenses()
        Dim conn As New OleDbConnection(NotificationModule.connectionString)
        Try
            Dim query As String = "SELECT Category, SUM(Amount) AS TotalSpent FROM Expenses WHERE Month(PurchaseDate) = Month(Date()) GROUP BY Category"
            Dim adapter As New OleDbDataAdapter(query, conn)
            Dim dt As New DataTable()
            adapter.Fill(dt)

            For Each row As DataRow In dt.Rows
                Dim category As String = row("Category").ToString()
                Dim totalSpent As Decimal = Convert.ToDecimal(row("TotalSpent"))

                Dim budgetLimit As Decimal = 500  ' Set a budget limit

                If totalSpent > budgetLimit Then
                    AddNotification($"Warning: Expenses for {category} exceeded the budget!", "Expense")
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine("Error checking expenses: " & ex.Message)
        End Try

    End Sub
    ' Function to check for overdue tasks and send a notification
    Private Sub CheckOverdueTasks()
        Dim conn As New OleDbConnection(NotificationModule.connectionString)
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
    'Private Sub SendNotification(message As String, category As String, userId As Integer)
    '    Dim conn As New OleDbConnection(NotificationModule.connectionString)
    '    Dim query As String = "INSERT INTO Notifications (Message, DateCreated, Category, UserId, IsRead) VALUES (@Message, @DateCreated, @Category, @UserId, IsRead)"

    '    Using connection As New OleDbConnection(connectionString)
    '        Dim command As New OleDbCommand(query, connection)
    '        command.Parameters.AddWithValue("@Message", message)
    '        command.Parameters.AddWithValue("@DateCreated", DateTime.Now)
    '        command.Parameters.AddWithValue("@Category", category)
    '        command.Parameters.AddWithValue("@UserId", userId)
    '        command.Parameters.AddWithValue("@IsRead", False)
    '        connection.Open()
    '        command.ExecuteNonQuery()
    '    End Using
    'End Sub

    ' Function to check for overdue chores and send a notification
    Private Sub CheckOverdueChores()
        Dim conn As New OleDbConnection(NotificationModule.connectionString)

        Try
            Dim query As String = "SELECT ChoreName FROM Chores WHERE DueDate < Date() AND Status <> 'Completed'"
            Dim adapter As New OleDbDataAdapter(query, conn)
            Dim dt As New DataTable()
            adapter.Fill(dt)

            For Each row As DataRow In dt.Rows
                AddNotification($"Chore '{row("ChoreName")}' is overdue!", "Chore")
            Next
        Catch ex As Exception
            Debug.WriteLine("Error checking overdue chores: " & ex.Message)
        End Try
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

    ''''''


End Class
