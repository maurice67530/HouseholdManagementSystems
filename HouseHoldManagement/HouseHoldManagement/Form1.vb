Public Class Form1
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Panel2_Paint(sender As Object, e As PaintEventArgs) Handles Panel2.Paint

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            Dim chore As New chores_() With {
           .Title =
           .AssignedTo =
           .Priority =
           .Status =
           .Frequency =
           .DueDate =
           .Recurring =
           .Description = }


            Dim tablename As String = "Chores"
            Dim Cmd As New OleDbCommand($"INSERT INTO {tablename} ([Title], [AssignedTo], [Priority], [Status], [Frequency], [DueDate], [Recurring], [Description]) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", conn)

            Cmd.Parameters.Clear()

            Cmd.Parameters.AddWithValue("@Title", chore.Title)
            Cmd.Parameters.AddWithValue("@AssignedTo", chore.AssignedTo)
            Cmd.Parameters.AddWithValue("@Priority", chore.Priority)
            Cmd.Parameters.AddWithValue("@Status", chore.Status)
            Cmd.Parameters.AddWithValue("@Frequency", chore.Frequency)
            Cmd.Parameters.AddWithValue("@DueDate", chore.DueDate)
            Cmd.Parameters.AddWithValue("@Recurring", chore.Recurring)
            Cmd.Parameters.AddWithValue("@Description", chore.Description)

            MsgBox("chores Information Addded!" & vbCrLf &
              "Title: " & chore.Title & vbCrLf &
              "AssignedTo:" & chore.Description & vbCrLf &
              "Priority: " & chore.Priority & vbCrLf &
              "Status : " & chore.Status & vbCrLf &
              "Frequency: " & chore.AssignedTo & vbCrLf &
                "Recurring: " & chore.Recurring & vbCrLf &
              "Description: " & chore.Description & vbCrLf &
              "DueDate: " & chore.DueDate & vbCrLf & vbCrLf, vbInformation, "Chores confirmation")

            MessageBox.Show("Chores Information saved to Database successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Cmd.ExecuteNonQuery()

        Catch ex As OleDbException
            Debug.WriteLine($"General error in button Save: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("Error saving control texts: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show($"Error Saving To database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show("Error saving Chores to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch ex As Exception
            MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Debug.WriteLine("Existing button Save")


    End Sub

End Class
