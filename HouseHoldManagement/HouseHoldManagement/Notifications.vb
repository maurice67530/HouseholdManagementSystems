Imports System.IO
Imports System.Data.OleDb
Imports System.Media
Public Class Notifications
    Private conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)

        If DataGridView1.SelectedRows.Count = 0 Then
            MsgBox("Please select a notification.", MsgBoxStyle.Exclamation, "No Selection")
            Return
        End If

        Conn.Open()

        For Each row As DataGridViewRow In DataGridView1.SelectedRows
            Dim notificationID As Integer = Convert.ToInt32(row.Cells("ID").Value)
            Dim currentStatus As String = row.Cells("IsRead").Value.ToString()
            Dim newStatus As String = If(currentStatus = "Yes", "No", "Yes") ' Toggle status
            Dim query As String = "UPDATE Notifications SET IsRead = @IsRead WHERE ID = @ID"

            Using cmd As New OleDbCommand(query, Conn)
                cmd.Parameters.AddWithValue("@IsRead", newStatus)
                cmd.Parameters.AddWithValue("@ID", notificationID)
                cmd.ExecuteNonQuery()
            End Using

            row.Cells("IsRead").Value = newStatus ' Update UI cell value
            row.DefaultCellStyle.ForeColor = If(newStatus = "Yes", Color.Black, Color.Red)
        Next

        Conn.Close()

        MsgBox("Selected notifications updated!", MsgBoxStyle.Information, "Updated")

        ' Update unread count in Label2
        UpdateUnreadCount()
    End Sub
    ' You can define this sub to update the label
    Sub UpdateUnreadCount()
        Try
            Using Conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                Conn.Open()
                Dim query As String = "SELECT COUNT(*) FROM Notifications WHERE IsRead = 'No'"
                Using cmd As New OleDbCommand(query, Conn)
                    Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                    Label2.Text = "Unread: " & count.ToString()
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine("Error counting unread notifications: " & ex.Message)
        End Try
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

        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
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

    Private Sub Notifications_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Start()
        Timer1.Interval = 2000

        ToolTip1.SetToolTip(Button1, "Mark As Read")
        ToolTip1.SetToolTip(Button2, "Clear Notification")
        ToolTip1.SetToolTip(Button3, "Refresh")
        LoadNotifications()

        'CheckInventoryAndChores()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub
    Private Sub CountUnreadNotifications()
        ' SQL query to count unread notifications
        Dim query As String = "SELECT COUNT(*) FROM Notifications WHERE IsRead = 'No'"

        ' Create a new connection to the database
        Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)

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

    'Private Sub CheckLowInventory()
    '    Dim localConn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
    '    Dim lowItems As New List(Of String) ' Store low stock items

    '    Try
    '        localConn.Open()
    '        Debug.WriteLine("Connection opened.")

    '        Using cmd As New OleDbCommand("SELECT ItemName, Quantity FROM Inventory WHERE LEN(Quantity) > 0 AND IsNumeric(Quantity) = True", localConn)
    '            Using reader As OleDbDataReader = cmd.ExecuteReader()
    '                While reader.Read()
    '                    Dim itemName As String = reader("ItemName").ToString()
    '                    Dim quantityString As String = reader("Quantity").ToString().Trim()
    '                    Dim quantity As Integer

    '                    If Integer.TryParse(quantityString, quantity) Then
    '                        If quantity <= 60 Then
    '                            Debug.WriteLine(itemName & ": Yes")
    '                            lowItems.Add(itemName & " (" & quantity & ")")
    '                            AddNotification(currentUser, itemName, quantity)
    '                        Else
    '                            Debug.WriteLine(itemName & ": No")
    '                        End If
    '                    Else
    '                        Debug.WriteLine("Invalid quantity for: " & itemName)
    '                    End If
    '                End While
    '            End Using
    '        End Using

    '        ' Show message only with low inventory items
    '        If lowItems.Count > 0 Then
    '            MessageBox.Show("Low Inventory Items:" & vbCrLf & String.Join(vbCrLf, lowItems), "Low Inventory", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '        End If

    '    Catch ex As Exception
    '        MessageBox.Show("Error checking inventory: " & ex.Message)
    '    Finally
    '        If localConn.State = ConnectionState.Open Then
    '            localConn.Close()
    '            Debug.WriteLine("Connection closed.")
    '        End If
    '    End Try
    'End Sub







    'Private Sub AddNotification(userID As String, itemName As String, quantity As Integer)
    '    ' Message to be sent
    '    Dim message As String = "Low inventory: " & itemName & " only has " & quantity.ToString()

    '    ' Convert Date.Now to a string (make sure the format is correct)
    '    Dim dateCreated As String = Date.Now.ToString("yyyy-MM-dd HH:mm:ss")
    '    Dim category As String = "Inventory"
    '    Dim isRead As String = "No"

    '    ' Check if the notification already exists in the DataGridView
    '    For Each row As DataGridViewRow In DataGridView1.Rows
    '        If row.Cells("Message").Value.ToString() = message Then
    '            ' If it already exists, show a message box with details of the existing notification
    '            Dim existingItem As String = row.Cells("Message").Value.ToString()
    '            'MessageBox.Show("The following notification is already in the DataGridView: " & vbCrLf & existingItem, "Notification Exists", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '            Exit Sub
    '        End If
    '    Next

    '    ' Create a new connection for inserting the notification
    '    Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)

    '    Try
    '        ' Open connection before executing the insert command
    '        conn.Open()

    '        ' Prepare the insert command
    '        Dim insertCmd As New OleDbCommand("INSERT INTO Notifications ([UserID], [Message], [DateCreated], [Category], [IsRead]) VALUES (?, ?, ?, ?, ?)", conn)

    '        ' Set parameter values explicitly for better clarity
    '        insertCmd.Parameters.AddWithValue("@UserID", userID)
    '        insertCmd.Parameters.AddWithValue("@Message", message)
    '        insertCmd.Parameters.AddWithValue("@DateCreated", dateCreated)
    '        insertCmd.Parameters.AddWithValue("@Category", category)
    '        insertCmd.Parameters.AddWithValue("@IsRead", isRead)

    '        ' Execute the insert command
    '        insertCmd.ExecuteNonQuery()

    '    Catch ex As OleDbException
    '        ' Handle database error
    '        MessageBox.Show("Database error occurred: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    Catch ex As Exception
    '        ' Handle unexpected error
    '        MessageBox.Show("An unexpected error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    Finally
    '        ' Ensure the connection is closed after the operation
    '        If conn.State = ConnectionState.Open Then
    '            conn.Close()
    '        End If
    '    End Try
    'End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Stop()
        CheckInventoryAndChores()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.RowIndex >= 0 Then
            Dim notificationId As Integer = DataGridView1.Rows(e.RowIndex).Cells("ID").Value
            Dim query As String = "UPDATE Notifications SET IsRead = True WHERE ID = " & notificationId
            ExecuteQuery(query)
        End If
        CheckInventoryAndChores()
        LoadNotifications()

    End Sub
    Private Sub LoadNotifications()
        Dim query As String = "SELECT * FROM Notifications ORDER BY DateCreated DESC"
        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString), cmd As New OleDbCommand(query, conn)
            Dim adapter As New OleDbDataAdapter(cmd)
            Dim table As New DataTable()
            adapter.Fill(table)
            DataGridView1.DataSource = table
        End Using
    End Sub
    Private Sub ExecuteQuery(query As String)
        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString), cmd As New OleDbCommand(query, conn)
            conn.Open()
            cmd.ExecuteNonQuery()
        End Using
        LoadNotifications()
        ShowNewNotifications()
    End Sub
    Private Sub ShowNewNotifications()
        Dim query As String = "SELECT Message FROM Notifications WHERE IsRead = False"
        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString), cmd As New OleDbCommand(query, conn)
            conn.Open()
            Using reader As OleDbDataReader = cmd.ExecuteReader()
                Dim messages As String = ""
                While reader.Read()
                    messages &= reader("Message").ToString() & vbCrLf
                End While
                If messages <> "" Then
                    MessageBox.Show("New Notifications:" & vbCrLf & messages, "Notifications", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        End Using
    End Sub

    Private Sub CheckInventoryAndChores()
        Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
        Dim currentUser As String = "System"
        Dim dateCreated As String = Date.Now.ToString("yyyy-MM-dd HH:mm:ss")
        Dim isRead As String = "No"

        Try
            conn.Open()

            ' Check Inventory (Quantity <= 60)
            Dim inventoryCmd As New OleDbCommand("SELECT ItemName, Quantity FROM Inventory WHERE LEN(Quantity) > 0 AND IsNumeric(Quantity) = True", conn)
            Dim inventoryReader As OleDbDataReader = inventoryCmd.ExecuteReader()
            While inventoryReader.Read()
                Dim itemName As String = inventoryReader("ItemName").ToString()
                Dim quantity As Integer
                If Integer.TryParse(inventoryReader("Quantity").ToString(), quantity) AndAlso quantity <= 60 Then
                    Dim message As String = "Low inventory: " & itemName & " only has " & quantity.ToString()
                    If Not NotificationExists(conn, message) Then
                        AddNotification(conn, currentUser, message, "Inventory", dateCreated, isRead)
                    End If
                End If
            End While
            inventoryReader.Close()

            ' Check Chores (Overdue)
            Dim choreCmd As New OleDbCommand("SELECT Title, DueDate FROM Chores", conn)
            Dim choreReader As OleDbDataReader = choreCmd.ExecuteReader()
            While choreReader.Read()
                Dim title As String = choreReader("Title").ToString()
                Dim dueDate As Date
                If Date.TryParse(choreReader("DueDate").ToString(), dueDate) AndAlso dueDate < Date.Today Then
                    Dim message As String = "Overdue chore: " & title & " was due on " & dueDate.ToShortDateString()
                    If Not NotificationExists(conn, message) Then
                        AddNotification(conn, currentUser, message, "Chore", dateCreated, isRead)
                    End If
                End If
            End While
            choreReader.Close()

            ' Optional Alert
            SystemSounds.Exclamation.Play()

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            If conn.State = ConnectionState.Open Then conn.Close()
        End Try
    End Sub

    Private Sub AddNotification(conn As OleDbConnection, userID As String, message As String, category As String, dateCreated As String, isRead As String)
        Try
            Dim insertQuery As String = "INSERT INTO Notifications ([UserID], [Message], [DateCreated], [Category], [IsRead]) VALUES (?, ?, ?, ?, ?)"
            Using cmd As New OleDbCommand(insertQuery, conn)
                cmd.Parameters.AddWithValue("?", userID)
                cmd.Parameters.AddWithValue("?", message)
                cmd.Parameters.AddWithValue("?", dateCreated)
                cmd.Parameters.AddWithValue("?", category)
                cmd.Parameters.AddWithValue("?", isRead)
                cmd.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            MessageBox.Show("Error saving notification: " & ex.Message)
        End Try
    End Sub

    Private Function NotificationExists(conn As OleDbConnection, message As String) As Boolean
        Dim checkQuery As String = "SELECT COUNT(*) FROM Notifications WHERE Message = ?"
        Using cmd As New OleDbCommand(checkQuery, conn)
            cmd.Parameters.AddWithValue("?", message)
            Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
            Return count > 0
        End Using
    End Function


End Class