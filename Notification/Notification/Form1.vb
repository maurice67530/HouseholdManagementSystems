Imports System.IO
Imports System.Data.OleDb
Public Class Form1

    'Dim tracker As New Form1()

    Dim conn As New OleDbConnection(Notifications.connectionString)

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

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Timer1.Interval = 60000 ' Check every 60 seconds

        Timer1.Start()

        LoadNotifications()

        CountUnreadNotifications()

        IbINotification.ForeColor = Color.White

        IbINotification.BackColor = Color.Red

        IbINotification.AutoSize = True

        IbINotification.Text = "New Notification!!"

        IbINotification.Visible = False

        Me.Controls.Add(IbINotification)


        ToolTip1.SetToolTip(Button1, "Mark As Read")

        ToolTip1.SetToolTip(btnClear, "Clear Notification")

        ToolTip1.SetToolTip(Button2, "Refresh")

    End Sub


    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

        Dim Conn As New OleDbConnection(Notifications.connectionString)

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

        'CountUnreadNotifications()


    End Sub

    Private Sub LoadNotifications()

        Dim query As String = "SELECT * FROM Notifications ORDER BY DateCreated DESC"

        Using conn As New OleDbConnection(Notifications.connectionString), cmd As New OleDbCommand(query, conn)

            Dim adapter As New OleDbDataAdapter(cmd)

            Dim table As New DataTable()

            adapter.Fill(table)

            DataGridView1.DataSource = table

        End Using

    End Sub

    Private Sub TrackExpenses()

        Dim query As String = "INSERT INTO Notifications (UserID, Message, DateCreated, Category, IsRead) " &
                   "SELECT 'Expense', 'Recent expense recorded: ' & Amount, Date(), False FROM Expenses WHERE Date >= Date()-7"

        ExecuteQuery(query)

    End Sub

    Private Sub TrackOverdueChores()

        Dim query As String = "INSERT INTO Notifications (UserID, Message, DateCreated, Category, IsRead) " &
                                  "SELECT 'Chores', 'Overdue chore: ' & ChoreName, DueDate, False FROM Chores WHERE Completed = False AND DueDate < Date()"

        ExecuteQuery(query)

    End Sub

    Private Sub TrackPendingTasks()

        Dim query As String = "INSERT INTO Notifications (UserID, Message, DateCreated, Category, IsRead) " &
                                  "SELECT 'Task', 'Pending task: ' & TaskName, DueDate, False FROM Tasks WHERE Completed = False"

        ExecuteQuery(query)

    End Sub



    Private Sub TrackLowInventory()

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

                                ' If the quantity is less than or equal to 60, send a notification

                                If quantity <= 60 Then

                                    AddNotification("System", itemName, quantity)

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

        Using conn As New OleDbConnection(Notifications.connectionString), cmd As New OleDbCommand(query, conn)

            conn.Open()

            cmd.ExecuteNonQuery()

        End Using

        LoadNotifications()

        ShowNewNotifications()

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

                End If

            End Using

        End Using

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

        If e.RowIndex >= 0 Then

            Dim notificationId As Integer = DataGridView1.Rows(e.RowIndex).Cells("ID").Value

            Dim query As String = "UPDATE Notifications SET IsRead = True WHERE ID = " & notificationId

            ExecuteQuery(query)

        End If

        TrackLowInventory()

        LoadNotifications()

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

        'CountUnreadNotifications()

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        TrackExpenses()

        TrackOverdueChores()

        TrackPendingTasks()

        TrackLowInventory()

        Timer1.Stop()

    End Sub

    Private Sub CountUnreadNotifications()

        Dim query As String = "SELECT COUNT(*) FROM Notifications WHERE IsRead = False"

        Using conn As New OleDbConnection(Notifications.connectionString), cmd As New OleDbCommand(query, conn)

            conn.Open()

            Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())

            Label2.Text = "Unread Notifications: " & count.ToString()

        End Using

    End Sub

End Class


