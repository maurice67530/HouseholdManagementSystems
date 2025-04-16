Imports System.Data.OleDb
Public Class Chores
    Public Property connect As New OleDbConnection(Ndivhuwo.connectionString)
    Public Property connn As New OleDbConnection(Masindi.connectionString)
    Public Property conn As New OleDbConnection(Murangi.connectionString)
    Private Sub Chores_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cmbpri.Items.AddRange(New String() {"Low", "Medium", "High"})
        cmbstatus.Items.AddRange(New String() {"Not started", "In progress", "Completed"})
        cmbfre.Items.AddRange(New String() {"Daily", "Weekly", "Monthly"})

        Dim tooltip As New ToolTip
        tooltip.SetToolTip(Button1, "Dashboard")
        tooltip.SetToolTip(Button2, "Mark All as Complete")
        tooltip.SetToolTip(Button3, "Refresh")
        tooltip.SetToolTip(Button4, "Delete")
        tooltip.SetToolTip(Button5, "Edit")
        tooltip.SetToolTip(Button6, "Submit")
        tooltip.SetToolTip(Button7, "Highlight")
        tooltip.SetToolTip(Button8, "Filter")
        tooltip.SetToolTip(Button9, "Sort")

        PopulateComboboxFromDatabase(cmbassi)

        loadChoresFromDatabase()
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        conn.Open()

        'Try
        Dim chores As New chores_() With {
           .Title = TXTtitle.Text,
           .AssignedTo = cmbassi.SelectedItem,
            .Priority = cmbpri.SelectedItem,
            .Status = cmbstatus.SelectedItem,
            .Frequency = cmbfre.SelectedItem,
            .DueDate = DateTimePicker1.Text,
            .Recurring = NumericUpDown1.Value,
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

    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Try
            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.Cells("DueDate").Value IsNot Nothing Then
                    Dim DueDate As DateTime = Convert.ToDateTime(row.Cells("DueDate").Value)
                    Dim Status As String = row.Cells("Status").Value.ToString

                    If DueDate < DateTime.Now AndAlso Status <> "Completed" Then
                        row.DefaultCellStyle.BackColor = Color.Gray

                    Else
                        row.DefaultCellStyle.BackColor = Color.WhiteSmoke


                    End If
                End If
            Next
            Dim InclompletedCount As Integer = 0

            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.Cells("Status").Value IsNot Nothing AndAlso row.Cells("Status").Value.ToString() <> "Completed " Then

                    InclompletedCount += 1

                End If
            Next
            Label10.Text = "Incompleted chores: " & InclompletedCount.ToString
        Catch ex As Exception
            MessageBox.Show("Error highligting overdue chores")
        End Try
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            ' Get the selected row  
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Assuming there is an "ID" column which is the primary key in the database  
            Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  

            ' Confirm deletion  

            ' Proceed with deletion  
            Try
                If DataGridView1.SelectedRows.Count > 0 Then
                    Debug.WriteLine("A row is selected for deletion")
                    Dim confirmationResults As DialogResult = MessageBox.Show("Are you sure you want o delete this chore?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                    If confirmationResults = DialogResult.Yes Then
                        Debug.WriteLine("User confirmed deletion.")

                    Else
                        Debug.WriteLine("User cancelled deletion.")
                    End If

                Else
                    Debug.WriteLine("No row selected, Exiting Button delete")

                    MessageBox.Show("Please Select chore to delete", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If

                Using conn As New OleDbConnection(Murangi.connectionString)
                    conn.Open()


                    ' Create the delete command  
                    Dim cmd As New OleDbCommand("DELETE FROM [Chores] WHERE [ID] = ?", conn)
                    cmd.Parameters.AddWithValue("@ID", ID) ' Primary key for matching record  

                    ' Execute the delete command  
                    Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                    If rowsAffected > 0 Then
                        MessageBox.Show("Task deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ' Optionally refresh DataGridView or reload from database  
                        loadChoresFromDatabase()
                    Else
                        MessageBox.Show("No chores deleted. Please check if the ID exists.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If

                End Using

            Catch ex As Exception
                MessageBox.Show($"An error occurred while deleting the expense: {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else

            MessageBox.Show("Please select chores to delete.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            loadChoresFromDatabase()
        End If
    End Sub
    Public Sub loadChoresFromDatabase()
        Try

            Debug.WriteLine("Datagridview: populated successfully.")

            Using conn As New OleDbConnection(Murangi.connectionString)
                conn.Open()

                Dim tableName As String = "Chores"

                Dim cmd As New OleDbCommand($"SELECT*FROM {tableName}", conn)

                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)

                DataGridView1.DataSource = dt

            End Using

        Catch ex As Exception
            Debug.WriteLine($"Datagridview: Failed to Populate {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error Loading Inventory data from database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        loadChoresFromDatabase()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        For Each row As DataGridViewRow In
            DataGridView1.SelectedRows
            row.Cells("Status").Value = "Completed"
        Next
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Close()
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try

            Debug.WriteLine("entering button update")

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If DataGridView1.SelectedRows.Count > 0 Then
                Debug.WriteLine("A row is selected for update")

            Else
                MessageBox.Show("Please select a Chore to update.", "update Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Debug.WriteLine("No row selected, exiting Button Update")
            End If
            Debug.WriteLine("Exiting Button Update")

            Dim Title As String = TXTtitle.Text
            Dim AssignedTo As String = cmbassi.SelectedItem
            Dim Priority As String = cmbpri.SelectedItem
            Dim Status As String = cmbstatus.SelectedItem
            Dim Frequency As String = cmbfre.SelectedItem
            Dim DueDate As String = DateTimePicker1.Value
            Dim Recurring As String = NumericUpDown1.Value
            Dim Description As String = txtdes.Text

            Using conn As New OleDbConnection(Murangi.connectionString)
                conn.Open()

                'Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                'Create an OleDbCommand to update the personnel data in the database  
                Dim cmd As New OleDbCommand("UPDATE Chores SET [Title] = ?, [AssignedTo] = ?, [Priority] = ?, [Status] = ?, [Frequency] = ?, [DueDate] = ?, [Recurring]= ?, [Description] = ? WHERE [ID] = ?", conn)

                'Set the parameter values from the UI controls  


                cmd.Parameters.AddWithValue("@Title", TXTtitle.Text)
                cmd.Parameters.AddWithValue("@AssignedTo", cmbassi.SelectedItem)
                cmd.Parameters.AddWithValue("@Priority", cmbpri.SelectedItem)
                cmd.Parameters.AddWithValue("@Status", cmbstatus.SelectedItem)
                cmd.Parameters.AddWithValue("@Frequency", cmbfre.SelectedItem)
                cmd.Parameters.AddWithValue("@DueDate", DateTimePicker1.Value)
                cmd.Parameters.AddWithValue("@Recurring", NumericUpDown1.Value)
                cmd.Parameters.AddWithValue("@Description", txtdes.Text)
                cmd.Parameters.AddWithValue("@ID", ID)
                MsgBox("Chores Updated Successfuly!", vbInformation, "Update Confirmation")

                cmd.ExecuteNonQuery()

            End Using
        Catch ex As FormatException
            Debug.WriteLine($"Format Error in Button Update: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error updating Tasks in database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show("Error saving inventory to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Debug.WriteLine($"unexpected error Button Update: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Unexpected error: {ex.Message}", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show("Error saving inventory to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        DataGridView1.Sort(DataGridView1.Columns("DueDate"), System.ComponentModel.ListSortDirection.Ascending)
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Try

            'Button1.Enabled = False

            If DataGridView1.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Debug.WriteLine("row selected on dgv")


                TXTtitle.Text = selectedRow.Cells("Title").Value.ToString()
                cmbassi.SelectedItem = selectedRow.Cells("AssignedTo").Value.ToString()
                cmbpri.SelectedItem = selectedRow.Cells("Priority").Value.ToString()
                cmbstatus.SelectedItem = selectedRow.Cells("Status").Value.ToString()
                cmbfre.SelectedItem = selectedRow.Cells("Frequency").Value.ToString()
                NumericUpDown1.Value = selectedRow.Cells("Recurring").Value.ToString()
                txtdes.Text = selectedRow.Cells("Description").Value.ToString()
                DateTimePicker1.Value = selectedRow.Cells("DueDate").Value.ToString()


            End If
            'Button1.Enabled = True
        Catch ex As Exception
            Debug.WriteLine("Data not selected: Error")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")

        End Try
    End Sub
    Public Sub PopulateComboboxFromDatabase(ByRef comboBox As ComboBox)
        Dim conn As New OleDbConnection(Murangi.connectionString)
        Try
            Debug.WriteLine("populate combobox successful")
            'open the database connection
            conn.Open()

            'retrieve the firstname and surname columns from the personaldetails tabel
            Dim query As String = "SELECT FirstName, LastName FROM PersonalDetails"
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            'bind the retrieved data to the combobox
            cmbassi.Items.Clear()
            While reader.Read()
                cmbassi.Items.Add($"{reader("FirstName")} {reader("LastName")}")
            End While

            'close the database
            reader.Close()

        Catch ex As Exception
            'handle any exeptions that may occur  
            Debug.WriteLine("failed to populate combobox")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error: {ex.StackTrace}")

        Finally
            'close the database connection
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
End Class