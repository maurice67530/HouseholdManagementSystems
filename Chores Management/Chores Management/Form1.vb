
Imports System.Data.OleDb
Public Class Form1
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'Try
        conn.Open()
        Dim chores As New chores_() With {
           .Title = TXTtitle.Text,
           .AssignedTo = cmbassi.SelectedItem,
            .Priority = cmbpri.SelectedItem,
            .Status = cmbstatus.SelectedItem,
            .Frequency = cmbfre.SelectedItem,
            .DueDate = DateTimePicker1.Text,
            .Recurring = NumericUpDown1.Text,
            .Description = txtdes.Text}


        Dim tablename As String = "Chores"
        Dim Cmd As New OleDbCommand($"INSERT INTO {tablename} ([Title], [AssignedTo], [Priority], [Status], [Frequency], [DueDate], [Recurring], [Description]) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", conn)

        Cmd.Parameters.Clear()

        Cmd.Parameters.AddWithValue("@Title", chores.Title)
        Cmd.Parameters.AddWithValue("@AssignedTo", chores.AssignedTo)
        Cmd.Parameters.AddWithValue("@Priority", chores.Priority)
        Cmd.Parameters.AddWithValue("@Status", chores.Status)
        Cmd.Parameters.AddWithValue("@Frequency", chores.Frequency)
        Cmd.Parameters.AddWithValue("@DueDate", chores.DueDate)
        Cmd.Parameters.AddWithValue("@Recurring", chores.Recurring)
        Cmd.Parameters.AddWithValue("@Description", chores.Description)

        MsgBox("chores Information Addded!" & vbCrLf &
              "Title: " & chores.Title & vbCrLf &
              "AssignedTo:" & chores.Description & vbCrLf &
              "Priority: " & chores.Priority & vbCrLf &
              "Status : " & chores.Status & vbCrLf &
              "Frequency: " & chores.AssignedTo & vbCrLf &
              "Recurring: " & chores.Recurring & vbCrLf &
              "Description: " & chores.Description & vbCrLf &
              "DueDate: " & chores.DueDate & vbCrLf & vbCrLf, vbInformation, "Chores confirmation")

        MessageBox.Show("Chores Information saved to Database successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Cmd.ExecuteNonQuery()

        'Catch ex As OleDbException
        '    Debug.WriteLine($"General error in button Save: {ex.Message}")
        '    Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
        '    MessageBox.Show("Error saving control texts: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    MessageBox.Show($"Error Saving To database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    MessageBox.Show("Error saving Chores to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        'Catch ex As Exception
        '    MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End Try
        'Debug.WriteLine("Existing button Save")
        conn.Close()


    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

    End Sub
End Class
