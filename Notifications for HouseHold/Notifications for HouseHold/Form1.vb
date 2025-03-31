Imports System.IO
Imports System.Data.OleDb



Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadNotifications()
        UpdateUnreadCount()
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
                Dim notificationID As Integer = reader("ID")
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


    Sub UpdateUnreadCount()
        Dim Conn As New OleDbConnection(Notifications.connectionString)

        Try
            ' Open Connection
            Conn.Open()

            Dim query As String = "SELECT COUNT(*) FROM Notifications WHERE IsRead=False"
            Dim cmd As New OleDbCommand(query, Conn)
            Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())

            Label1.Text = "Unread: " & count

        Catch ex As Exception
            MessageBox.Show("Error updating unread count: " & ex.Message)

        Finally
            ' Close Connection
            If Conn.State = ConnectionState.Open Then
                Conn.Close()
            End If
        End Try
    End Sub

    Sub AddNotification(userID As Integer, message As String)
        Dim Conn As New OleDbConnection

        Dim query As String = "INSERT INTO Notifications (UserID, Message, DateCreated, IsRead) VALUES (@UserID, @Message, @DateCreated, False)"
        Dim cmd As New OleDbCommand(query, Conn)
        cmd.Parameters.AddWithValue("@UserID", userID)
        cmd.Parameters.AddWithValue("@Message", message)
        cmd.Parameters.AddWithValue("@DateCreated", DateTime.Now)
        cmd.ExecuteNonQuery()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        UpdateUnreadCount()
    End Sub
End Class
