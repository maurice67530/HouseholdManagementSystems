Imports System.Data.OleDb
Public Class Task_Management
    Private Sub Task_Management_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim tooltip As New ToolTip
        tooltip.SetToolTip(Button1, "Submit")
        tooltip.SetToolTip(Button4, "Refresh")
        tooltip.SetToolTip(Button3, "Delete")
        tooltip.SetToolTip(Button2, "Edit")
        tooltip.SetToolTip(Button7, "Dashboard")
        tooltip.SetToolTip(Button5, "Filter")
        tooltip.SetToolTip(Button6, "Sort")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Debug.WriteLine("entering btnsave")

            If TextBox2.Text = " " Then
                MsgBox("please enter a value")
                TextBox2.Focus()

            End If
            If ComboBox1.Text = "" Then
                MsgBox("please enter a value")
                ComboBox1.Focus()

            End If
            If DateTimePicker1.Text = "" Then
                MsgBox("please enter a value")
                DateTimePicker1.Focus()

            End If
            Using conn As New OleDbConnection(connectionString)
                conn.Open()

                Dim tableName As String = "Tasks"
                Dim cmd As New OleDbCommand("INSERT INTO Tasks ([Title],[Description],[DueDate],[Priority],[Status],[AssignedTo]) VALUES (@Title, @Description, @DueDate, @Priority, @Status, @AssignedTo)", conn)

                Dim Task As New DailyTask() With {
             .Title = TextBox1.Text,
              .Description = TextBox2.Text,
            .DueDate = DateTimePicker1.Text,
            .Priority = ComboBox1.Text,
            .Status = ComboBox2.Text,
            .AssignedTo = ComboBox3.Text}


                'parameters

                cmd.Parameters.AddWithValue("@Title", Task.Title)
                cmd.Parameters.AddWithValue("@Description", Task.Description)
                cmd.Parameters.AddWithValue("@DueDate", Task.DueDate)
                cmd.Parameters.AddWithValue("@Priority", Task.Priority)
                cmd.Parameters.AddWithValue("@Status", Task.Status)
                cmd.Parameters.AddWithValue("@AssignedTo", Task.AssignedTo)

                cmd.ExecuteNonQuery()

                MsgBox("tasks Added!" & vbCrLf &
                   "Title: " & Task.Title.ToString() & vbCrLf &
                   "Description: " & Task.Description & vbCrLf &
                   "DueDate: " & Task.DueDate & vbCrLf &
                   "Priority: " & Task.Priority & vbCrLf &
                   "Status: " & Task.Status & vbCrLf &
                   "Assignedto: " & Task.AssignedTo, vbInformation, "inventory Confirmation")


            End Using
        Catch ex As OleDbException
            Debug.WriteLine($"database error: {ex.Message}")
            MessageBox.Show("error saving test to database. please check the connection.", "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch ex As Exception
            Debug.WriteLine($"general error in btnsave:{ex.Message}")
            MessageBox.Show("An unexpected error occurred.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

        Debug.WriteLine("Exit btnsave")
    End Sub

End Class