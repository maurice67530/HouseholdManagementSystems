Imports System.IO
Imports System.Data.OleDb
Public Class Form1
    Dim conn As New OleDbConnection(Notifications.connectionString)
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadNotifications()
        UpdateUnreadCount()
        CheckOverdueChores()
        CheckOverdueTasks()
        CheckInventory()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Conn As New OleDbConnection(Notifications.connectionString)


        Try
            ' Open the connection
            Conn.Open()

            ' Execute the delete query
            Dim query As String = "DELETE FROM Notifications"
            Dim cmd As New OleDbCommand(query, Conn)
            cmd.ExecuteNonQuery()

            ' Reload the list
            LoadNotifications()

        Catch ex As Exception
            MessageBox.Show("Error deleting notifications: " & ex.Message)

        Finally
            ' Close the connection
            If Conn.State = ConnectionState.Open Then
                Conn.Close()
            End If
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Conn As New OleDbConnection(Notifications.connectionString)

        If ListBox1.SelectedIndex <> -1 Then
            Dim selectedText As String = ListBox1.SelectedItem.ToString()
            Dim notificationID As Integer = Convert.ToInt32(selectedText.Split(" - ")(0)) ' Extract ID

            Dim query As String = "UPDATE Notifications SET IsRead=True WHERE NotificationID=@ID"
            Dim cmd As New OleDbCommand(query, Conn)
            cmd.Parameters.AddWithValue("@ID", notificationID)
            cmd.ExecuteNonQuery()

            LoadNotifications()
        Else
            MessageBox.Show("Select a notification to mark as read.")
        End If
    End Sub

    Sub LoadNotifications()
        Dim Conn As New OleDbConnection(Notifications.connectionString)

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
                Dim notificationID As Integer = reader("NotificationID")
                Dim message As String = reader("Message")
                ListBox1.Items.Add(notificationID & " - " & message) ' Display ID & Message
            End While

            reader.Close()
            UpdateUnreadCount()

        Catch ex As Exception
            MessageBox.Show("Error loading notifications: " & ex.Message)

        Finally
            ' Close Connection
            If Conn.State = ConnectionState.Open Then
                Conn.Close()
            End If
        End Try
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        UpdateUnreadCount()
    End Sub

    Sub UpdateUnreadCount()
        Dim Conn As New OleDbConnection(Notifications.connectionString)

        Try
            ' Open Connection
            Conn.Open()

            Dim query As String = "SELECT COUNT(*) FROM Notifications WHERE IsRead=False"
            Dim cmd As New OleDbCommand(query, Conn)
            Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())

            Label2.Text = "Unread: " & count

        Catch ex As Exception
            MessageBox.Show("Error updating unread count: " & ex.Message)

        Finally
            ' Close Connection
            If Conn.State = ConnectionState.Open Then
                Conn.Close()
            End If
        End Try
    End Sub

    ' Function to check inventory quantity and send notification if below 5
    Private Sub CheckInventory()
        Dim conn As New OleDbConnection(Notifications.connectionString)
        Dim query As String = "SELECT Item, Quantity FROM Inventory WHERE Quantity < 5"

        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(query, connection)
            connection.Open()

            Dim reader As OleDbDataReader = command.ExecuteReader()

            ' Check if any items are below the threshold of 5
            If reader.HasRows Then
                While reader.Read()
                    Dim item As String = reader("Item").ToString()
                    Dim quantity As Integer = Convert.ToInt32(reader("Quantity"))

                    ' Trigger a notification for low stock
                    If quantity < 5 Then
                        ' You can use a simple MessageBox for a notification
                        MessageBox.Show($"Low Inventory Alert: {item} is below 5 units. Current stock: {quantity}", "Low Inventory", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                        ' Alternatively, you can use a custom notification control or system tray notification.
                        ' NotifyIcon1.ShowBalloonTip(3000, "Low Inventory", $"{item} is below 5 units!", ToolTipIcon.Warning)
                    End If
                End While
            End If

            reader.Close()
        End Using
    End Sub

    ' Function to check for overdue chores and send a notification
    Private Sub CheckOverdueChores()
        Dim conn As New OleDbConnection(Notifications.connectionString)
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

    ' Function to check for overdue tasks and send a notification
    Private Sub CheckOverdueTasks()
        Dim conn As New OleDbConnection(Notifications.connectionString)
        Dim query As String = "SELECT TaskName, DueDate, AssignedTo FROM Tasks WHERE DueDate < @CurrentDate AND Status <> 'Completed'"

        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(query, connection)
            command.Parameters.AddWithValue("@CurrentDate", DateTime.Now) ' Current Date

            connection.Open()

            Dim reader As OleDbDataReader = command.ExecuteReader()

            ' Check if there are any overdue tasks
            If reader.HasRows Then
                While reader.Read()
                    Dim taskName As String = reader("TaskName").ToString()
                    Dim dueDate As DateTime = Convert.ToDateTime(reader("DueDate"))
                    Dim assignedTo As Integer = Convert.ToInt32(reader("AssignedTo"))

                    ' Trigger overdue notification
                    Dim message As String = $"Overdue Task: {taskName} was due on {dueDate.ToShortDateString()}."
                    SendNotification(message, "Task", assignedTo)
                End While
            End If

            reader.Close()
        End Using
    End Sub

    ' Function to add a notification to the database
    Private Sub SendNotification(message As String, category As String, userId As Integer)
        Dim connectionString As String = Notifications.connectionString
        Dim query As String = "INSERT INTO Notifications (Message, DateCreated, Category, UserId) VALUES (@Message, @DateCreated, @Category, @UserId)"

        Using connection As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(query, connection)
            command.Parameters.AddWithValue("@Message", message)
            command.Parameters.AddWithValue("@DateCreated", DateTime.Now)
            command.Parameters.AddWithValue("@Category", category)
            command.Parameters.AddWithValue("@UserId", userId)

            connection.Open()
            command.ExecuteNonQuery()
        End Using
    End Sub

End Class
