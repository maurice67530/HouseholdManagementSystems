Imports System.Data.OleDb
Public Class Form1
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Debug.WriteLine("Entering button delete")
        ' Check if there are any selected rows in the DataGridView for expenses  
        If DataGridView1.SelectedRows.Count > 0 Then

            ' Get the selected row  
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Assuming there is an "ID" column which is the primary key in the database  
            Dim InventoryID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  
            'Dim DeletedBy As String

            ' Confirm deletion  
            Dim confirmationResult As DialogResult = MessageBox.Show("Are you sure you want to delete this?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If confirmationResult = DialogResult.Yes Then
                Debug.WriteLine("User confirmation deletion.")
                ' Proceed with deletion  
                Try
                    Debug.WriteLine("Format errors in button delete")
                    Debug.WriteLine("Deleting data: Data delected")
                    Debug.WriteLine("Stack Trace: {ex.StackTrace}")
                    Using conn As New OleDbConnection(InventoryModule.connectionString)
                        conn.Open()

                        ' Create the delete command  
                        Dim cmd As New OleDbCommand("DELETE FROM [Inventory] WHERE [ID] = ?", conn)
                        cmd.Parameters.AddWithValue("@ID", InventoryID) ' Primary key for matching record  

                        ' Execute the delete command  
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show("Inventory deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ' Optionally refresh DataGridView or reload from database  
                            ' PopulateDataGridView()
                            LoadInventorydataFromDatabase()
                        Else
                            MessageBox.Show("No Inventory was deleted. Please check if the ID exists.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End Using
                    LoadInventorydataFromDatabase()
                Catch ex As Exception
                    Debug.WriteLine("Failed to delete data")
                    Debug.Write($"Stack Trace: {ex.StackTrace}")
                    MessageBox.Show($"An error occurred while deleting the Inventory: {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            Debug.WriteLine("User cancelled deletion")
            MessageBox.Show("Please select Inventory to delete.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

        Debug.WriteLine("Exiting button delete")
    End Sub
    Public Sub LoadInventorydataFromDatabase()
        Try
            '  Dim dataTable As DataTable = HouseHold.GetData("SELECT * FROM Expense")
            ' DataGridView1.DataSource = DataTable
            Debug.WriteLine("Populate Datagridview: Datagridview populated successfully.")
            Using conn As New OleDbConnection(InventoryModule.connectionString)
                conn.Open()

                Dim tableName As String = "Inventory"

                Dim cmd As New OleDbCommand($"SELECT*FROM {tableName}", conn)

                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)
                DataGridView1.DataSource = dt
            End Using
        Catch ex As Exception
            Debug.WriteLine("Failed to populate Gatagridview")
            Debug.Write($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error Loading expense data from database: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Debug.WriteLine("Entering button update click")
        If DataGridView1.SelectedRows.Count = 0 Then
            Debug.WriteLine("User confirmed update")
            MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Debug.WriteLine($"Format error in button update:")
            Debug.WriteLine("Updating data: data updated")
            Dim ItemName As String = TextBox1.Text
            Dim Description As String = TextBox2.Text
            Dim Quantity As String = TextBox3.Text
            Dim Category As String = ComboBox1.SelectedItem.ToString
            Dim ReorderLevel As String = TextBox4.Text
            Dim PricePerUnit As String = TextBox5.Text
            Dim DateAdded As String = DateTimePicker1.Value
            Dim ExpiryDate As String = DateTimePicker1.Value
            Dim Unit As String = ComboBox2.SelectedItem.ToString

            Using conn As New OleDbConnection(InventoryModule.connectionString)
                conn.Open()

                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Create an OleDbCommand to update the personnel data in the database  
                Dim cmd As New OleDbCommand("UPDATE [Inventory] SET [ItemName] = ?, [Description] = ?, [Quantity] = ?, [Category] = ?, [ReorderLevel] = ?, [PricePerUnit] = ?, [DateAdded] = ?,  [ExpiryDate] = ?, [Unit] = ? WHERE [ID] = ?", conn)

                ' Set the parameter values from the UI controls  

                cmd.Parameters.AddWithValue("@ItemName", ItemName)
                cmd.Parameters.AddWithValue("@Description", Description)
                cmd.Parameters.AddWithValue("@Quantity", Quantity)
                cmd.Parameters.AddWithValue("@Category", Category)
                cmd.Parameters.AddWithValue("@ReorderLevel", ReorderLevel)
                cmd.Parameters.AddWithValue("@PricePerUnit", PricePerUnit)
                cmd.Parameters.AddWithValue("@DateAdded", DateAdded)
                cmd.Parameters.AddWithValue("@ExpiryDate", ExpiryDate)
                cmd.Parameters.AddWithValue("@Unit", Unit)
                cmd.Parameters.AddWithValue("@ID", ID)

                cmd.ExecuteNonQuery()

                MsgBox("Inventory Items Updated Successfuly!", vbInformation, "Update Confirmation")
                LoadInventorydataFromDatabase()
                '  InventoryModule.ClearControls(Me)
            End Using
        Catch ex As OleDbException
            Debug.WriteLine("User cancelled update")
            Debug.WriteLine("Unexpected error in button update")
            Debug.Write($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error updating inventory in database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Debug.WriteLine("Exiting button update")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadInventorydataFromDatabase()
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class