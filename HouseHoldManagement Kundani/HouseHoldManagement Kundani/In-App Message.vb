Imports System.Data.OleDb
Public Class In_App_Message

    Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles txtReply.TextChanged

    End Sub

    Private Sub In_App_Message_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'ShowLastViewedMessages()
        NotifyIcon1.ShowBalloonTip(5000) ' 5 seconds


        LoadNewOverdueChores()
        LoadNewHighExpenses()
        UpdateLastViewed() ' Mark as read after loading
    End Sub

    Private Sub ShowLastViewedMessages()
        Dim cmd As New OleDbCommand("SELECT ViewType, LastViewed FROM NotificationStatus", conn)

        conn.Open()
        Dim rdr As OleDbDataReader = cmd.ExecuteReader()
        Dim notifyMsg As String = ""

        While rdr.Read()
            Dim viewType As String = rdr("ViewType").ToString()
            Dim lastViewed As DateTime = CDate(rdr("LastViewed"))
            Dim message As String = $"Last viewed {viewType}: {lastViewed}"

            ' Show MessageBox
            MessageBox.Show(message, "Notification Info", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' Append for NotifyIcon
            notifyMsg &= message & vbCrLf
        End While

        conn.Close()

        ' Show NotifyIcon message (balloon)
        If notifyMsg <> "" Then
            NotifyIcon1.BalloonTipTitle = "In-App Notifications"
            NotifyIcon1.BalloonTipText = notifyMsg
            NotifyIcon1.Visible = True
            NotifyIcon1.ShowBalloonTip(5000) ' 5 seconds
        End If
    End Sub





    Private Sub LoadNewOverdueChores()
        ListBox1.Items.Clear()
        Dim lastViewed As DateTime = GetLastViewed("Chores")
        Dim cmd As New OleDbCommand("SELECT Title, DueDate FROM Chores WHERE DueDate > ? AND DueDate > ?", conn)
        cmd.Parameters.AddWithValue("?", Date.Today)
        cmd.Parameters.AddWithValue("?", lastViewed)

        conn.Open()
        Dim rdr As OleDbDataReader = cmd.ExecuteReader()
        While rdr.Read()
            ListBox1.Items.Add("Overdue Chores: " & rdr("Title") & " (Due " & CDate(rdr("DueDate")).ToShortDateString() & ")")
        End While
        conn.Close()
    End Sub

    Private Sub LoadNewHighExpenses()
        ListBox2.Items.Clear()
        Dim lastViewed As DateTime = GetLastViewed("Expense")
        Dim threshold As Decimal = 1000D
        Dim cmd As New OleDbCommand("SELECT Category, Amount, DateOfexpenses FROM Expense WHERE Amount > ? AND DateOfexpenses > ?", conn)
        cmd.Parameters.AddWithValue("?", threshold)
        cmd.Parameters.AddWithValue("?", lastViewed)

        conn.Open()
        Dim rdr As OleDbDataReader = cmd.ExecuteReader()
        While rdr.Read()
            ListBox2.Items.Add("High Expense: " & rdr("Category") & " - ₱" & FormatNumber(rdr("Amount"), 2))
        End While
        conn.Close()
    End Sub
    Private Function GetLastViewed(viewType As String) As DateTime
        Dim cmd As New OleDbCommand("SELECT LastViewed FROM NotificationStatus WHERE ViewType = ?", conn)
        cmd.Parameters.AddWithValue("?", viewType)

        conn.Open()
        Dim result As Object = cmd.ExecuteScalar()
        conn.Close()

        If result IsNot Nothing Then
            Return CDate(result)
        Else
            ' Insert initial record if not exists
            cmd = New OleDbCommand("INSERT INTO NotificationStatus (ViewType, LastViewed) VALUES (?, ?)", conn)
            cmd.Parameters.AddWithValue("?", viewType)
            cmd.Parameters.AddWithValue("?", DateTime.MinValue)
            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
            Return DateTime.MinValue
        End If
    End Function

    Private Sub UpdateLastViewed()
        Dim nowTime As DateTime = DateTime.Now
        Dim views() As String = {"Chores", "Expenses"}

        For Each view In views
            Dim cmd As New OleDbCommand("UPDATE NotificationStatus SET LastViewed = ? WHERE ViewType = ?", conn)
            cmd.Parameters.AddWithValue("?", nowTime)
            cmd.Parameters.AddWithValue("?", view)

            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
        Next
    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub
    'Public Sub PopulateComboboxFromDatabase(ByRef comboBox As ComboBox)
    '    Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
    '    Try
    '        Debug.WriteLine("populate combobox successful")
    '        'open the database connection
    '        conn.Open()

    '        'retrieve the firstname and surname columns from the personaldetails tabel
    '        Dim query As String = "SELECT FirstName, LastName FROM Users"
    '        Dim cmd As New OleDbCommand(query, conn)
    '        Dim reader As OleDbDataReader = cmd.ExecuteReader()

    '        'bind the retrieved data to the combobox
    '        ComboBox1.Items.Clear()
    '        While reader.Read()
    '            ComboBox1.Items.Add($"{reader("FullNames")} {reader("UserName")}")
    '        End While

    '        'close the database
    '        reader.Close()

    '    Catch ex As Exception
    '        'handle any exeptions that may occur  
    '        Debug.WriteLine("failed to populate combobox")
    '        Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
    '        MessageBox.Show($"Error: {ex.StackTrace}")

    '    Finally
    '        'close the database connection
    '        If conn.State = ConnectionState.Open Then
    '            conn.Close()
    '        End If
    '    End Try
    'End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        LoadNotifications()
    End Sub

    Private Sub chkShowAll_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowAll.CheckedChanged
        LoadNotifications()
    End Sub
    Private Sub LoadNotifications()
        LoadNewOverdueChores()
        LoadNewHighExpenses()
        If Not chkShowAll.Checked Then
            UpdateLastViewed()
        End If
    End Sub

    Private Sub btnSendReply_Click(sender As Object, e As EventArgs) Handles btnSendReply.Click
        If originalMessage = "" Or txtReply.Text.Trim() = "" Then
            MessageBox.Show("Please select a message and write a reply.", "Missing Info", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim con As New OleDbConnection(Murangi.connectionString)
        Dim cmd As New OleDbCommand("INSERT INTO NotificationReplies (NotificationType, OriginalMessage, ReplyMessage, ReplyDate) VALUES (?, ?, ?, ?)", con)
        'cmd.Parameters.AddWithValue("?", currentUserName)
        cmd.Parameters.AddWithValue("?", selectedMessageType)
        cmd.Parameters.AddWithValue("?", originalMessage)
        cmd.Parameters.AddWithValue("?", txtReply.Text.Trim())
        cmd.Parameters.AddWithValue("?", DateTime.Now)

        con.Open()
        cmd.ExecuteNonQuery()
        con.Close()

        MessageBox.Show("Reply sent and saved successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        txtReply.Clear()
    End Sub



    Private selectedMessageType As String = ""
    Private originalMessage As String = ""

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        If ListBox1.SelectedItem IsNot Nothing Then
            selectedMessageType = "Chores"
            originalMessage = ListBox1.SelectedItem.ToString()
            txtReply.Text = "" ' Clear old reply
        End If
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        If ListBox2.SelectedItem IsNot Nothing Then
            selectedMessageType = "Expenses"
            originalMessage = ListBox2.SelectedItem.ToString()
            txtReply.Text = ""
        End If
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick

    End Sub
End Class





























