Imports System.IO
Imports System.Data.OleDb

Public Class Form1





    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Run all checks And then load notifications
        CheckDueTasks()
        CheckLowInventory()
        CheckHighExpenses()
        LoadNotifications()
        CountUnreadNotifications()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            If ListBox1.SelectedIndex <> -1 Then
                Dim selectedText As String = ListBox1.SelectedItem.ToString()
                Dim notificationID As Integer = Integer.Parse(selectedText.Split(" - ")(0))

                Dim query As String = "UPDATE Notifications SET IsRead = True WHERE ID = @id"
                Dim cmd As New OleDbCommand(query, conn)
                cmd.Parameters.AddWithValue("@id", notificationID)
                conn.Open()
                cmd.ExecuteNonQuery()
                conn.Close()

                LoadNotifications()
            Else
                MessageBox.Show("Please select a notification to mark as read.")
            End If
        Catch ex As Exception
            Debug.WriteLine("Error marking notification as read: " & ex.Message)
        End Try
    End Sub

    ' Loads unread notifications into ListBox1
    Private Sub LoadNotifications()
        Try
            ListBox1.Items.Clear()
            Dim query As String = "SELECT ID, Message FROM Notifications WHERE IsRead = False"
            Dim cmd As New OleDbCommand(query, conn)
            conn.Open()
            Dim reader As OleDbDataReader = cmd.ExecuteReader()
            While reader.Read()
                ' Add each notification as "ID - Message"
                ListBox1.Items.Add(reader("ID").ToString() & " - " & reader("Message").ToString())
            End While
            conn.Close()
        Catch ex As Exception
            Debug.WriteLine("Error loading notifications: " & ex.Message)
        End Try
    End Sub



    ' Function to add a new notification
    Public Sub AddNotification(message As String)
        Try
            Dim query As String = "INSERT INTO Notifications (Message, DateCreated, IsRead) VALUES (@msg, NOW(), False)"
            Dim cmd As New OleDbCommand(query, conn)
            cmd.Parameters.AddWithValue("@msg", message)
            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
        Catch ex As Exception
            Debug.WriteLine("Error adding notification: " & ex.Message)
        End Try
    End Sub

    ' Check for tasks that are due soon (within one day)
    Public Sub CheckDueTasks()
        Try
            Dim query As String = "SELECT TaskName FROM Chores WHERE DueDate <= Date() + 1"
            Dim adapter As New OleDbDataAdapter(query, conn)
            Dim dt As New DataTable()
            adapter.Fill(dt)
            For Each row As DataRow In dt.Rows
                AddNotification("Task '" & row("TaskName") & "' is due soon!")
            Next
        Catch ex As Exception
            Debug.WriteLine("Error checking due tasks: " & ex.Message)
        End Try
    End Sub

    ' Check for inventory items that are running low (less than 5)
    Public Sub CheckLowInventory()
        Try
            Dim query As String = "SELECT ItemName FROM Inventory WHERE Total < 5"
            Dim adapter As New OleDbDataAdapter(query, conn)
            Dim dt As New DataTable()
            adapter.Fill(dt)
            For Each row As DataRow In dt.Rows
                AddNotification("Low stock alert: " & row("ItemName") & " is running low!")
            Next
        Catch ex As Exception
            Debug.WriteLine("Error checking low inventory: " & ex.Message)
        End Try
    End Sub

    ' Check if expenses exceed 80% of a 6000 budget (i.e., 4800)
    Public Sub CheckHighExpenses()
        Try
            Dim totalExpense As Double = 0
            Dim query As String = "SELECT SUM(Amount) FROM Expenses"
            Dim cmd As New OleDbCommand(query, conn)
            conn.Open()
            Dim result = cmd.ExecuteScalar()
            conn.Close()

            If result IsNot DBNull.Value Then
                totalExpense = Convert.ToDouble(result)
            End If

            If totalExpense >= 4800 Then
                AddNotification("Warning: Expenses have exceeded 80% of the budget!")
            End If
        Catch ex As Exception
            Debug.WriteLine("Error checking high expenses: " & ex.Message)
        End Try
    End Sub

    Private Sub CountUnreadNotifications()
        Try
            Dim query As String = "SELECT COUNT(*) FROM Notifications WHERE IsRead = False"
            Dim cmd As New OleDbCommand(query, conn)
            conn.Open()
            Dim unreadCount As Integer = Convert.ToInt32(cmd.ExecuteScalar())
            conn.Close()

            ' Update the label with the unread count
            Label1.Text = "Unread Notifications: " & unreadCount
        Catch ex As Exception
            Debug.WriteLine("Error counting unread notifications: " & ex.Message)
        End Try
    End Sub






End Class
