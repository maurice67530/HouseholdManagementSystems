Imports System.IO
Imports System.Data.OleDb
Public Class Form1

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
        CountUnreadNotifications()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadNotifications()
        CountUnreadNotifications()



        IbINotification.ForeColor = Color.White
        IbINotification.BackColor = Color.Red
        IbINotification.AutoSize = True
        IbINotification.Text = "New Notification!!"
        IbINotification.Visible = False
        Me.Controls.Add(IbINotification)

        Timer1.Interval = 500
        Timer1.Enabled = False

        ToolTip1.SetToolTip(Button1, "Mark As Read")
        ToolTip1.SetToolTip(Button2, "Clear Notification")
        ToolTip1.SetToolTip(Button3, "Refresh")
    End Sub

    Private Sub LoadNotifications()
        Dim query As String = "SELECT * FROM Notifications ORDER BY DateCreated DESC"
        Using conn As New OleDbConnection(NotificationModule.connectionString), cmd As New OleDbCommand(query, conn)
            Dim adapter As New OleDbDataAdapter(cmd)
            Dim table As New DataTable()
            adapter.Fill(table)
            DataGridView1.DataSource = table
        End Using
    End Sub

    'Private Sub TrackExpenses()
    '    Dim query As String = "INSERT INTO Notifications (Category, Message, DateCreated, IsRead) " &
    '                      "SELECT 'Expenses', 'Recent expense recorded: ' & Amount, Date(), False FROM Expenses WHERE Date >= Date()-7"
    '    ExecuteQuery(query)
    'End Sub

    Private Sub TrackOverdueChores()
        Try
            Dim query As String = "INSERT INTO Notifications (Category, Message, DateCreated, IsRead) " &
                          "SELECT 'Chores', 'Overdue chore: ' & Title, DueDate, False FROM Chores WHERE Completed = False AND DueDate < Date()"
            ExecuteQuery(query)
        Catch ex As Exception
            MessageBox.Show("Error tracking chores: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TrackPendingTasks()
        Try
            Dim query As String = "INSERT INTO Notifications (Category, Message, DateCreated, IsRead) " &
                          "SELECT 'Tasks', 'Pending task: ' & Title, DueDate, False FROM Tasks WHERE Completed = False"
            ExecuteQuery(query)

        Catch ex As Exception
            MessageBox.Show("Error tracking tasks: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TrackLowInventory()
        Try
            Dim query As String = "INSERT INTO Notifications (Category, Message, DateCreated, IsRead) " &
                          "SELECT 'Inventory', 'Low inventory: ' & ItemName, DateAdded, False FROM Inventory WHERE Quantity < 65"
            ExecuteQuery(query)
        Catch ex As Exception
            MessageBox.Show("Error tracking chores: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ExecuteQuery(query As String)
        Using conn As New OleDbConnection(NotificationModule.connectionString), cmd As New OleDbCommand(query, conn)
            conn.Open()
            cmd.ExecuteNonQuery()
        End Using
        LoadNotifications()
        ShowNewNotifications()
    End Sub

    Private Sub ShowNewNotifications()

        Dim query As String = "SELECT Message FROM Notifications WHERE IsRead = False"

        Using conn As New OleDbConnection(NotificationModule.connectionString), cmd As New OleDbCommand(query, conn)
            conn.Open()

            Using reader As OleDbDataReader = cmd.ExecuteReader()
                Dim messages As String = ""

                While reader.Read()
                    messages &= reader("Message").ToString() & vbCrLf
                End While

                'If messages <> "" Then
                '    MessageBox.Show("New Notifications:" & vbCrLf & messages, "Notifications", MessageBoxButtons.OK, MessageBoxIcon.Information)

                '    ' If there are unread notifications, start blinking the label
                '    Timer1.Start()
                '    Label3.Text = "You have new notifications!"
                '    Label3.Visible = True
                'Else
                '    Timer1.Stop()
                '    Label3.Visible = False
                'End If
            End Using
        End Using
    End Sub




    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.RowIndex >= 0 Then
            Dim notificationId As Integer = DataGridView1.Rows(e.RowIndex).Cells("ID").Value
            Dim query As String = "UPDATE Notifications SET IsRead = True WHERE ID = " & notificationId
            ExecuteQuery(query)
            CountUnreadNotifications()

            'Timer1.Stop()
            'Label3.Visible = False

        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'TrackExpenses()
        'TrackOverdueChores()
        'TrackPendingTasks()
        'TrackLowInventory()
        'Label3.Visible = Not Label3.Visible
    End Sub

    Private Sub CountUnreadNotifications()
        Dim query As String = "SELECT COUNT(*) FROM Notifications WHERE IsRead = False"
        Using conn As New OleDbConnection(NotificationModule.connectionString), cmd As New OleDbCommand(query, conn)
            conn.Open()
            Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
            Label2.Text = "Unread Notifications: " & count.ToString()
        End Using
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'TrackExpenses()
        TrackOverdueChores()
        TrackPendingTasks()
        TrackLowInventory()
    End Sub

    Private Sub AddNotification(Category As String, message As String)
        Dim query As String = "INSERT INTO Notifications (Category, Message, DateCreated, IsRead) VALUES (@Category, @Message, Date(), False)"
        Using conn As New OleDbConnection(NotificationModule.connectionString), cmd As New OleDbCommand(query, conn)
            cmd.Parameters.AddWithValue("@Category", Category)
            cmd.Parameters.AddWithValue("@Message", message)
            conn.Open()
            cmd.ExecuteNonQuery()
        End Using
        LoadNotifications()
    End Sub


End Class

