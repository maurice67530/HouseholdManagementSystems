Imports System.IO
Imports System.Data.OleDb
Public Class Form1
    'Dim tracker As New Form1()
    Dim conn As New OleDbConnection(NotificationModule.connectionString)
    Private IbINotification As New Label
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.SelectedRows.Count = 0 Then
            MsgBox("Please select a notification to delete.", MsgBoxStyle.Exclamation, "No Selection")
            Return
        End If

        Dim confirmDelete As DialogResult = MessageBox.Show("Are you sure you want to delete the selected notifications?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If confirmDelete = DialogResult.No Then
            Return
        End If

        Using conn As New OleDbConnection(NotificationModule.connectionString)
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

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Interval = 2000 ' Check every 60 seconds
        Timer1.Start()

        IbINotification.Left = Label2.Left
        IbINotification.Top = Label2.Top - IbINotification.Height
        IbINotification.ForeColor = Color.White
        IbINotification.BackColor = Color.Red
        IbINotification.AutoSize = True
        IbINotification.Text = "New Notification!!"
        IbINotification.Visible = False

        Me.Controls.Add(IbINotification)


        ToolTip1.SetToolTip(Button1, "Mark As Read")
        ToolTip1.SetToolTip(Button2, "Clear Notification")
        ToolTip1.SetToolTip(Button3, "Refresh")

        LoadNotifications()

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Conn As New OleDbConnection(NotificationModule.connectionString)

        If DataGridView1.SelectedRows.Count = 0 Then
            MsgBox("Please select a notification to mark as read.", MsgBoxStyle.Exclamation, "No Selection")
            Return
        End If

        Conn.Open()

        For Each row As DataGridViewRow In DataGridView1.SelectedRows
            Dim notificationID As Integer = Convert.ToInt32(row.Cells("ID").Value)
            Dim query As String = "UPDATE Notifications SET IsRead = 'Yes' WHERE ID = @ID"

            Using cmd As New OleDbCommand(query, Conn)
                cmd.Parameters.AddWithValue("@ID", notificationID)
                cmd.ExecuteNonQuery()
            End Using

            row.DefaultCellStyle.ForeColor = Color.Black ' Change UI for read status
            row.Cells("IsRead").Value = "Yes"
        Next

        Conn.Close()
        MsgBox("Selected notifications marked as read!", MsgBoxStyle.Information, "Updated")

        ' Refresh unread count after marking as read
        CountUnreadNotifications()


    End Sub
    Private Sub LoadNotifications()
        Dim query As String = "SELECT * FROM Notifications"
        Using conn As New OleDbConnection(NotificationModule.connectionString), cmd As New OleDbCommand(query, conn)
            Dim adapter As New OleDbDataAdapter(cmd)
            Dim table As New DataTable()
            adapter.Fill(table)
            DataGridView1.DataSource = table
        End Using
    End Sub

    'Private Sub TrackExpenses()
    '    Dim query As String = "INSERT INTO Notifications (UserID, Message, DateCreated, Category, IsRead) " &
    '                              "SELECT 'Expense', 'Recent expense recorded: ' & Amount, DateOfexpense, False FROM Expenses WHERE Date >= Date()-7"
    '    ExecuteQuery(query)
    'End Sub

    'Private Sub TrackOverdueChores()
    '    Dim query As String = "INSERT INTO Notifications (UserID, Message, DateCreated, Category, IsRead) " &
    '                              "SELECT 'Chores', 'Overdue chore: ' & ChoreName, DueDate, False FROM Chores WHERE Completed = False AND DueDate < Date()"
    '    ExecuteQuery(query)
    'End Sub

    'Private Sub TrackPendingTasks()
    '    Dim query As String = "INSERT INTO Notifications (UserID, Message, DateCreated, Category, IsRead) " &
    '                              "SELECT 'Task', 'Pending task: ' & TaskName, DueDate, False FROM Tasks WHERE Completed = False"
    '    ExecuteQuery(query)
    'End Sub

    Private Sub TrackLowInventory()
        ' Create a new connection for this method to avoid issues with global connections.
        Dim localConn As New OleDbConnection(NotificationModule.connectionString)

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
                                ' If the quantity is less than or equal to 60, send a notification
                                If quantity <= 60 Then
                                    AddNotification("System", itemName, quantity)

                                    'Display a messagebox
                                    MessageBox.Show("Low Inventory: " & itemName & " Only has " & quantity.ToString() & " Left.", "Low Inventory Alert", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                End If
                            Else
                                ' Handle case where Quantity is not numeric
                                MessageBox.Show("Non-numeric quantity found for item: " & itemName)
                            End If
                        End If
                    End While
                End Using
            End Using

        Catch ex As Exception
            ' Show error message in case of failure
            MessageBox.Show("Error checking inventory: " & ex.Message)
        Finally
            ' Ensure the connection is closed here
            If localConn.State = ConnectionState.Open Then
                localConn.Close()
                Debug.WriteLine("Connection closed.")
            End If
        End Try
    End Sub
    Private Sub AddNotification(userID As String, itemName As String, quantity As Integer)
        Dim message As String = "Low inventory: " & itemName & " only has " & quantity.ToString()

        ' Convert Date.Now to a string (make sure the format is correct)
        Dim dateCreated As String = Date.Now.ToString("yyyy-MM-dd HH:mm:ss")

        Dim category As String = "Inventory"
        Dim isRead As String = "No"
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells("Message").Value.ToString = message Then
                Dim existingItem As String = row.Cells("Message").Value.ToString
                Exit Sub
            End If
        Next
        ' Open connection before executing the insert command
        conn.Open()

        Dim insertCmd As New OleDbCommand("INSERT INTO Notifications ([UserID], [Message], [DateCreated], [Category], [IsRead]) VALUES (?, ?, ?, ?, ?)", conn)

        ' Set parameter values explicitly for better clarity
        insertCmd.Parameters.AddWithValue("@UserID", userID)
        insertCmd.Parameters.AddWithValue("@Message", message)
        insertCmd.Parameters.AddWithValue("@DateCreated", dateCreated)
        insertCmd.Parameters.AddWithValue("@Category", category)
        insertCmd.Parameters.AddWithValue("@IsRead", isRead)

        Try
            insertCmd.ExecuteNonQuery()
        Catch ex As OleDbException
            MessageBox.Show("Database error occurred: " & ex.Message)
        Catch ex As Exception
            MessageBox.Show("An unexpected error occurred: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub ExecuteQuery(query As String)
        Using conn As New OleDbConnection(NotificationModule.connectionString), cmd As New OleDbCommand(query, conn)
            conn.Open()
            cmd.ExecuteNonQuery()
        End Using
        LoadNotifications()

    End Sub

    'Private Sub ShowNewNotifications()
    '    Dim query As String = "SELECT Message FROM Notifications WHERE IsRead = False"
    '    Using conn As New OleDbConnection(NotificationModule.connectionString), cmd As New OleDbCommand(query, conn)
    '        conn.Open()
    '        Using reader As OleDbDataReader = cmd.ExecuteReader()
    '            Dim messages As String = ""
    '            While reader.Read()
    '                messages &= reader("Message").ToString() & vbCrLf
    '            End While
    '            If messages <> "" Then
    '                MessageBox.Show("New Notifications:" & vbCrLf & messages, "Notifications", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '            End If
    '        End Using
    '    End Using
    'End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.RowIndex >= 0 Then
            Dim notificationId As Integer = DataGridView1.Rows(e.RowIndex).Cells("ID").Value
            Dim query As String = "UPDATE Notifications SET IsRead = True WHERE ID = " & notificationId
            ExecuteQuery(query)
        End If

    End Sub

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

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Timer1.Stop()
        IbINotification.Visible = Not IbINotification.Visible
        If IbINotification.ForeColor = Color.Red Then
            IbINotification.ForeColor = Color.Black
        Else
            IbINotification.ForeColor = Color.Red
        End If

        TrackLowInventory()
        CountUnreadNotifications()

    End Sub
    Private Sub CountUnreadNotifications()
        ' SQL query to count unread notifications
        Dim query As String = "SELECT COUNT(*) FROM Notifications WHERE IsRead = 'No'"

        ' Create a new connection to the database
        Dim conn As New OleDbConnection(NotificationModule.connectionString)

        Try
            ' Open the connection
            conn.Open()

            ' Prepare the command
            Dim cmd As New OleDbCommand(query, conn)

            ' Execute the command and get the count
            Dim unreadCount As Integer = Convert.ToInt32(cmd.ExecuteScalar())

            ' Update Label1 to display the number of unread notifications
            Label2.Text = "Unread Notifications: " & unreadCount.ToString()

            'Show notification is there are unread notifications
            If unreadCount > 0 Then
                MessageBox.Show("You have" & unreadCount.ToString & " Unread notification(s).", "New Notification", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        LoadNotifications()
        CountUnreadNotifications()
    End Sub
End Class




