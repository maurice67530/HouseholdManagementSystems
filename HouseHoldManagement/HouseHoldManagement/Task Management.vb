﻿Imports System.Data.OleDb

Public Class Task_Management
    Private ReadOnly Priority As String
    Public Property Status As String
    Public Property Tasks As Object

    Private Sub Task_Management_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim tooltip As New ToolTip
        tooltip.SetToolTip(Button1, "Submit")
        tooltip.SetToolTip(Button4, "Refresh")
        tooltip.SetToolTip(Button3, "Delete")
        tooltip.SetToolTip(Button2, "Edit")
        tooltip.SetToolTip(Button8, "Dashboard")
        tooltip.SetToolTip(Button5, "Filter")
        tooltip.SetToolTip(Button6, "Sort")

        ComboBox1.Items.AddRange(New String() {"Low", "Medium", "High"})
        ComboBox2.Items.AddRange(New String() {"Not started", "In progress", "Completed"})
    End Sub
    Private Sub ComboBox3_click(sender As Object, e As EventArgs) Handles ComboBox3.Click
        PopulateComboboxFromDatabase(ComboBox3)
    End Sub
    Public Sub PopulateComboboxFromDatabase(ByRef comboBox As ComboBox)
        Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
        Try
            Debug.WriteLine("populating combobox from database successfully!")
            ' 1. Open the database connection  
            conn.Open()

            ' 2. Retrieve the FirstName and LastName columns from the PersonalDetails table  
            Dim query As String = "SELECT FirstName, LastName FROM Person"
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            ' 3. Bind the retrieved data to the combobox  
            comboBox.Items.Clear()
            While reader.Read()
                comboBox.Items.Add($"{reader("FirstName")} {reader("LastName")}")
            End While

            ' 4. Close the database connection  
            reader.Close()

        Catch ex As Exception
            Debug.WriteLine($"form loaded unsuccessful")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
        Finally
            ' Close the database connection  
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Check if there are any selected rows in the DataGridView for personal  
        If DataGridView1.SelectedRows.Count > 0 Then
            ' Get the selected row  
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Assuming there is an "ID" column which is the primary key in the database  
            Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  

            ' Confirm deletion  
            Dim confirmationResult As DialogResult = MessageBox.Show("Are you sure you want to delete this e?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If confirmationResult = DialogResult.Yes Then
                ' Proceed with deletion  
                Try
                    Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                        conn.Open()

                        ' Create the delete command  
                        Dim cmd As New OleDbCommand("DELETE FROM [Tasks] WHERE [ID] = ?", conn)
                        cmd.Parameters.AddWithValue("@ID", ID) ' Primary key for matching record  

                        ' Execute the delete command  
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show("task deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ' Optionally refresh DataGridView or reload from database  
                            'PopulateDataGridView1()
                        Else
                            MessageBox.Show("No Task was deleted. Please check if the ID exists.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End Using

                Catch ex As Exception
                    MessageBox.Show($"An error occurred while deleting the Task:  {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            MessageBox.Show("Please select an Task to delete.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
        Dim TasksTable As New DataTable()
        Try
            conn.Open()
            Dim query As String = "SELECT * FROM Tasks WHERE 1=1"

            ' Only add conditions if filters are selected  
            If Not String.IsNullOrEmpty(Status) Then
                query &= " AND Status = @Status"
            End If

            If Not String.IsNullOrEmpty(Priority) Then
                query &= " AND Priority = @Priority"
            End If

            Dim command As New OleDb.OleDbCommand(query, conn)

            ' Only add parameters if filters are selected  
            If Not String.IsNullOrEmpty(Status) Then
                command.Parameters.AddWithValue("@Status", Status)
            End If

            If Not String.IsNullOrEmpty(Priority) Then
                command.Parameters.AddWithValue("@Priority", Priority)
            End If

            Dim adapter As New OleDb.OleDbDataAdapter(command)
            adapter.Fill(TasksTable)
            Tasks.DataGridView1.DataSource = TasksTable
        Catch ex As Exception
            MsgBox("Error filtering tasks: " & ex.Message, MsgBoxStyle.Critical, "Database Error")
        Finally
            conn.Close()
        End Try

        'get selected values from combobox
        Dim selectedStatus As String = If(ComboBox2.SelectedItem IsNot Nothing, ComboBox2.SelectedItem.ToString(), "")
        Dim selectedPriority As String = If(ComboBox1.SelectedItem IsNot Nothing, ComboBox1.SelectedItem.ToString(), "")

    End Sub
End Class