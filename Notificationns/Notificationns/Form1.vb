﻿Imports System.IO
Imports System.Data.OleDb

Public Class Form1



    Dim Conn As New OleDbConnection(Notifications.connectionString)
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim ToolTip1 As New ToolTip

        ToolTip1.SetToolTip(Button1, "UnRead Notification")
        ToolTip1.SetToolTip(Button2, "Clear Notification")

        Timer1.Start()
        Timer1.Interval = 6000
        LoadNotifications()
        CountUnreadNotifications()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.SelectedRows.Count = 0 Then
            MsgBox("Please select a notification to delete.", MsgBoxStyle.Exclamation, "No Selection")
            Return
        End If

        Dim confirmDelete As DialogResult = MessageBox.Show("Are you sure you want to delete the selected notifications?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If confirmDelete = DialogResult.No Then
            Return
        End If

        Using conn As New OleDbConnection(Notifications.connectionString)
            conn.Open()
            For Each row As DataGridViewRow In DataGridView1.SelectedRows
                Dim notificationID As Integer = Convert.ToInt32(row.Cells("ID").Value)
                Dim query As String = "DELETE FROM Notifications WHERE ID = @ID"
                Using cmd As New OleDbCommand(query, conn)
                    cmd.Parameters.AddWithValue("@ID", notificationID)
                    cmd.ExecuteNonQuery()
                End Using
                DataGridView1.Rows.Remove(row)
            Next
        End Using

        MsgBox("Selected notifications cleared!", MsgBoxStyle.Information, "Deleted")
        'CountUnreadNotifications()
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Conn As New OleDbConnection(Notifications.connectionString)

        If DataGridView1.SelectedRows.Count = 0 Then
            MsgBox("Please select a notification to mark as read.", MsgBoxStyle.Exclamation, "No Selection")
            Return
        End If

        Conn.Open()

        For Each row As DataGridViewRow In DataGridView1.SelectedRows
            Dim notificationID As Integer = Convert.ToInt32(row.Cells("ID").Value)

            ' Update IsRead to 'Yes' when marked as read
            Dim query As String = "UPDATE Notifications SET IsRead = 'Yes' WHERE ID = @ID"

            Using cmd As New OleDbCommand(query, Conn)
                cmd.Parameters.AddWithValue("@ID", notificationID)
                cmd.ExecuteNonQuery()
            End Using

            ' Change UI for read status (set text color to black for read notifications)
            row.DefaultCellStyle.ForeColor = Color.Black

            ' Update the IsRead cell value to 'Yes' in the DataGridView for UI consistency
            row.Cells("IsRead").Value = "Yes" ' Ensure "IsRead" is the correct column name for this

        Next

        Conn.Close()

        MsgBox("Selected notifications marked as read!", MsgBoxStyle.Information, "Updated")

        ' Refresh unread count after marking as read
        'CountUnreadNotifications()

    End Sub




    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ' Stop the timer after it ticks once
        Timer1.Stop()
        'LoadNotifications()
        ' Call the function to check for low inventory
        CheckLowInventory()
        LoadNotifications()
    End Sub


    Private Sub CheckLowInventory()
        ' Create a new connection for this method to avoid issues with global connections.
        Dim localConn As New OleDbConnection(Notifications.connectionString)

        Try
            ' Open the connection.
            localConn.Open()
            Debug.WriteLine("Connection opened.")

            ' Use 'Using' to manage the command and reader
            Using cmd As New OleDbCommand("SELECT ItemName, Quantity FROM Inventory WHERE LEN(Quantity) > 0 AND IsNumeric(Quantity) = True", localConn)
                Using reader As OleDbDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim itemName As String = reader("ItemName").ToString()

                        ' Get the Quantity value as a string
                        If Not IsDBNull(reader("Quantity")) Then
                            Dim quantityString As String = reader("Quantity").ToString().Trim()

                            ' Attempt to convert the Quantity string to Integer if it's numeric
                            Dim quantity As Integer
                            If Integer.TryParse(quantityString, quantity) Then
                                ' If the quantity is less than 60, send a notification and show a message box
                                If quantity < 60 Then
                                    ' Send the notification
                                    AddNotification("System", itemName, quantity)

                                    ' Display a message box
                                    MessageBox.Show("Low inventory: " & itemName & " only has " & quantity.ToString() & " left.", "Low Inventory Alert", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                End If
                            Else
                                ' Handle case where Quantity is not numeric
                                MessageBox.Show("Non-numeric quantity found for item: " & itemName, "Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            End If
                        End If
                    End While
                End Using
            End Using

        Catch ex As Exception
            ' Show error message in case of failure
            MessageBox.Show("Error checking inventory: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Ensure the connection is closed here
            If localConn.State = ConnectionState.Open Then
                localConn.Close()
                Debug.WriteLine("Connection closed.")
            End If
        End Try
    End Sub


    Private Sub AddNotification(userID As String, itemName As String, quantity As Integer)
        ' Message to be sent
        Dim message As String = "Low inventory: " & itemName & " only has " & quantity.ToString()

        ' Convert Date.Now to a string (make sure the format is correct)
        Dim dateCreated As String = Date.Now.ToString("yyyy-MM-dd HH:mm:ss")
        Dim category As String = "Inventory"
        Dim isRead As String = "No"

        ' Check if the notification already exists in the DataGridView
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells("Message").Value.ToString() = message Then
                ' If it already exists, show a message box with details of the existing notification
                Dim existingItem As String = row.Cells("Message").Value.ToString()
                'MessageBox.Show("The following notification is already in the DataGridView: " & vbCrLf & existingItem, "Notification Exists", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Next

        ' Create a new connection for inserting the notification
        Dim conn As New OleDbConnection(Notifications.connectionString)

        Try
            ' Open connection before executing the insert command
            conn.Open()

            ' Prepare the insert command
            Dim insertCmd As New OleDbCommand("INSERT INTO Notifications ([UserID], [Message], [DateCreated], [Category], [IsRead]) VALUES (?, ?, ?, ?, ?)", conn)

            ' Set parameter values explicitly for better clarity
            insertCmd.Parameters.AddWithValue("@UserID", userID)
            insertCmd.Parameters.AddWithValue("@Message", message)
            insertCmd.Parameters.AddWithValue("@DateCreated", dateCreated)
            insertCmd.Parameters.AddWithValue("@Category", category)
            insertCmd.Parameters.AddWithValue("@IsRead", isRead)

            ' Execute the insert command
            insertCmd.ExecuteNonQuery()

        Catch ex As OleDbException
            ' Handle database error
            MessageBox.Show("Database error occurred: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            ' Handle unexpected error
            MessageBox.Show("An unexpected error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Ensure the connection is closed after the operation
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub







    Private Sub LoadNotifications()
        Dim query As String = "SELECT * FROM Notifications"
        Using conn As New OleDbConnection(Notifications.connectionString), cmd As New OleDbCommand(query, conn)
            Dim adapter As New OleDbDataAdapter(cmd)
            Dim table As New DataTable()
            adapter.Fill(table)
            DataGridView1.DataSource = table
        End Using
    End Sub

    Private Sub ExecuteQuery(query As String)
        Using conn As New OleDbConnection(Notifications.connectionString), cmd As New OleDbCommand(query, conn)
            conn.Open()
            cmd.ExecuteNonQuery()
        End Using
        LoadNotifications()

    End Sub
    Private Sub ShowNewNotifications()

        Dim query As String = "SELECT Message FROM Notifications WHERE IsRead = False"

        Using conn As New OleDbConnection(Notifications.connectionString), cmd As New OleDbCommand(query, conn)
            conn.Open()

            Using reader As OleDbDataReader = cmd.ExecuteReader()
                Dim messages As String = ""

                While reader.Read()
                    messages &= reader("Message").ToString() & vbCrLf
                End While

                If messages <> "" Then
                    MessageBox.Show("New Notifications:" & vbCrLf & messages, "Notifications", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    ' If there are unread notifications, start blinking the label
                    Timer1.Start()
                    Label1.Text = "You have new notifications!"
                    Label1.Visible = True
                Else
                    Timer1.Stop()
                    Label1.Visible = False
                End If
            End Using
        End Using
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.RowIndex >= 0 Then
            Dim notificationId As Integer = DataGridView1.Rows(e.RowIndex).Cells("ID").Value
            Dim query As String = "UPDATE Notifications SET IsRead = True WHERE ID = " & notificationId
            ExecuteQuery(query)
            'CountUnreadNotifications()

            Timer1.Stop()
            Label1.Visible = False
        End If
    End Sub


    Private Sub CountUnreadNotifications()
        ' SQL query to count unread notifications
        Dim query As String = "SELECT COUNT(*) FROM Notifications WHERE IsRead = 'No'"

        ' Create a new connection to the database
        Dim conn As New OleDbConnection(Notifications.connectionString)

        Try
            ' Open the connection
            conn.Open()

            ' Prepare the command
            Dim cmd As New OleDbCommand(query, conn)

            ' Execute the command and get the count
            Dim unreadCount As Integer = Convert.ToInt32(cmd.ExecuteScalar())

            ' Update Label1 to display the number of unread notifications
            Label1.Text = "Unread Notifications: " & unreadCount.ToString()

        Catch ex As OleDbException
            ' Handle database error
            MessageBox.Show("Database error occurred: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            ' Handle unexpected error
            MessageBox.Show("An unexpected error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Ensure the connection is closed after the operation
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub




End Class
