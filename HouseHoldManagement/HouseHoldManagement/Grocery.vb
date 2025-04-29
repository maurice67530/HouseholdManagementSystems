Imports System.Data.OleDb
Public Class Grocery
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Rinae\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"
    Private Sub Grocery_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        LoadGroceryItemDataFromDatabase()
        '  notify()

        'Set tooltips for buttons
        ToolTip1.SetToolTip(btnSave, "Save")
        ToolTip1.SetToolTip(btnEdit, "Update")
        ToolTip1.SetToolTip(btnDelete, "Delete")
        ToolTip1.SetToolTip(btnSort, "Sort")
        ToolTip1.SetToolTip(btnFilter, "Filter")
        ToolTip1.SetToolTip(btnRefresh, "Refresh")
        ToolTip1.SetToolTip(btnHighlight, "Highlight")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try

            Dim Item As New Groceryy() With {
            .ItemName = TextBox1.Text,
            .Quantity = TextBox2.Text,
            .Category = TextBox3.Text,
            .Unit = TextBox4.Text,
            .PricePerUnit = TextBox6.Text,
            .Period = ComboBox2.SelectedItem,
            .PurchaseDate = DateTimePicker1.Text,
            .ExpiryDate = DateTimePicker2.Text
             }

            Using conn As New OleDbConnection(Rinae.connectionString)
                conn.Open()

                Dim cmd As New OleDbCommand($"INSERT INTO GroceryItems ([ItemName], [Quantity], [Category], [Unit], [Period], [PricePerUnit], [PurchaseDate], [ExpiryDate]) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", conn)
                cmd.Parameters.AddWithValue("@ItemName", TextBox1.Text)
                cmd.Parameters.AddWithValue("@Quantity", TextBox2.Text)
                cmd.Parameters.AddWithValue("@Category", TextBox3.Text)
                cmd.Parameters.AddWithValue("@Unit", TextBox4.Text)
                cmd.Parameters.AddWithValue("@Period", ComboBox2.SelectedItem.ToString)
                cmd.Parameters.AddWithValue("@PricePerUnit", TextBox6.Text)
                cmd.Parameters.AddWithValue("@PurchaseDate", DateTimePicker1.Value.ToString)
                cmd.Parameters.AddWithValue("@ExpiryDate", DateTimePicker2.Value.ToString)

                'Execute the SQL command to insert the data
                cmd.ExecuteNonQuery()
                LoadGroceryItemDataFromDatabase()
                conn.Close()

            End Using

            'Display a Confirmation Message  
            MessageBox.Show("Names details added successfully")

        Catch ex As OleDbException
            Debug.WriteLine($"Database error in btnSubmit_Click: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("Error saving GroceryItem to database. Please check the connection.", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch ex As Exception
            Debug.WriteLine($"General error in btnSubmit_Click: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("An unexpected error occurred.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
        LoadGroceryItemDataFromDatabase()
        Debug.WriteLine("Exiting btnSubmit_Click ")
    End Sub
    Public Sub LoadGroceryItemDataFromDatabase()
        Try

            Debug.WriteLine("DataGridView loaded succesful")

            Using conn As New OleDbConnection(Rinae.connectionString)
                conn.Open()

                ' Update the table name if necessary  
                Dim tableName As String = "GroceryItems"

                ' Create an OleDbCommand to select the data from the database  
                Dim cmd As New OleDbCommand($"SELECT * FROM {tableName}", conn)

                ' Create a DataAdapter and fill a DataTable  
                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable()
                da.Fill(dt)

                ' Bind the DataTable to the DataGridView  
                DataGridView1.DataSource = dt
                conn.Close()
            End Using

        Catch ex As OleDbException
            Debug.WriteLine("Data loading fail")
            MessageBox.Show("Error loading GroceryItem data from database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch ex As Exception

            Debug.WriteLine($"Error in PopulateDataGridView: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("An error occurred while loading data into the grid.", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
        Debug.WriteLine("Exiting  PopulateDataGridView")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
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
            Dim Quantity As String = TextBox2.Text
            Dim Category As String = TextBox3.Text
            Dim Unit As String = TextBox4.Text
            Dim PricePerUnit As String = TextBox6.Text
            Dim Period As String = ComboBox2.SelectedItem.ToString
            Dim PurchaseDate As String = DateTimePicker1.Value
            Dim ExpiryDate As String = DateTimePicker2.Value

            Using conn As New OleDbConnection(Rinae.connectionString)
                conn.Open()

                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Create an OleDbCommand to update the personnel data in the database  
                Dim cmd As New OleDbCommand("UPDATE [GroceryItems] SET [ItemName] = ?, [Quantity] = ?, [Category] = ?, [Unit] = ?, [PricePerUnit] = ?, [Period] = ?, [PurchaseDate] = ?, [ExpiryDate] = ? WHERE [ID] = ?", conn)

                ' Set the parameter values from the UI controls  

                'cmd.Parameters.AddWithValue("@PhotoID", PhotoID.textbox1.Text)
                cmd.Parameters.AddWithValue("@ItemName", ItemName)
                cmd.Parameters.AddWithValue("@Quantity", Quantity)
                cmd.Parameters.AddWithValue("@Category", Category)
                cmd.Parameters.AddWithValue("@Unit", Unit)
                cmd.Parameters.AddWithValue("@PricePerUnit", PricePerUnit)
                cmd.Parameters.AddWithValue("@Period", Period)
                cmd.Parameters.AddWithValue("@PurchaseDate", PurchaseDate)
                cmd.Parameters.AddWithValue("@ExpiryDate", ExpiryDate)

                cmd.Parameters.AddWithValue("@ID", ID)

                cmd.ExecuteNonQuery()

                MsgBox("Grocery Items Updated Successfuly!", vbInformation, "Update Confirmation")
                LoadGroceryItemDataFromDatabase()
                '   Cruwza.ClearControls(Me)
            End Using
        Catch ex As OleDbException
            Debug.WriteLine("User cancelled update")
            Debug.WriteLine("Unexpected error in button update")
            Debug.Write($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error updating grocery in database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Debug.WriteLine("Exiting button update")
    End Sub
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged

        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
            ' Load the data from the selected row into UI controls  
            TextBox1.Text = selectedRow.Cells("ItemName").Value.ToString()
            TextBox2.Text = selectedRow.Cells("Quantity").Value.ToString()
            TextBox3.Text = selectedRow.Cells("Category").Value.ToString()
            TextBox4.Text = selectedRow.Cells("Unit").Value.ToString()
            ComboBox2.SelectedItem = selectedRow.Cells("Period").Value.ToString()
            TextBox6.Text = selectedRow.Cells("PricePerUnit").Value.ToString()
            DateTimePicker1.Value = selectedRow.Cells("PurchaseDate").Value.ToString()
            DateTimePicker2.Value = selectedRow.Cells("ExpiryDate").Value.ToString()
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles btnDelete.Click

        ' Check if there are any selected rows in the DataGridView for expenses  
        If DataGridView1.SelectedRows.Count > 0 Then

            ' Get the selected row  
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Assuming there is an "ID" column which is the primary key in the database  
            Dim ChoreID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name 

            ' Confirm deletion  
            Dim confirmationResult As DialogResult = MessageBox.Show("Are you sure you want to delete this chore?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If confirmationResult = DialogResult.Yes Then
                Debug.WriteLine("User confirmation deletion.")
                ' Proceed with deletion  
                Try
                    Debug.WriteLine("Format errors in button delete")
                    Debug.WriteLine("Deleting data: Data delected")
                    Debug.WriteLine("Stack Trace: {ex.StackTrace}")
                    Using conn As New OleDbConnection(Rinae.connectionString)
                        conn.Open()

                        Dim cmd As New OleDbCommand("DELETE FROM [GroceryItem] WHERE [ID] = ?", conn)
                        cmd.Parameters.AddWithValue("@ID", ChoreID) ' Primary key for matching record  

                        ' Execute the delete command  
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show("item deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ' Optionally refresh DataGridView or reload from database  
                            'PopulateDataGridView()
                        Else
                            Debug.WriteLine("User cancelled deletion")
                            MessageBox.Show("No item was deleted. Please check if the ID exists.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End Using
                    LoadGroceryItemDataFromDatabase()

                Catch ex As Exception
                    Debug.WriteLine("Failed to delete data")
                    Debug.Write($"Stack Trace: {ex.StackTrace}")
                    'MessageBox.Show($"An error occurred while deleting the expense: {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            Debug.WriteLine("No row selected, existing btnDelete_Click")
            MessageBox.Show("Please select an chore to delete.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        Debug.WriteLine("Exiting btnDelete_Click")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnHighlight.Click
        HighlightExpiredItems()
    End Sub
    Private Sub HighlightExpiredItems()

        ' Assuming you have a DataGridView named "dataGridView1"
        ' and the expiration date column is named "ExpiryDate"

        ' Loop through each row in the DataGridView
        For Each row As DataGridViewRow In DataGridView1.Rows
            ' Check if the row is not a new row (in case the user is entering data)
            If Not row.IsNewRow Then
                ' Get the expiration date value from the respective column
                Dim ExpiryDate As DateTime
                If DateTime.TryParse(row.Cells("ExpiryDate").Value.ToString(), ExpiryDate) Then
                    ' Compare the expiration date with the current date
                    If ExpiryDate < DateTime.Now Then
                        ' Highlight the row by changing its background color
                        row.DefaultCellStyle.BackColor = Color.Red
                        row.DefaultCellStyle.ForeColor = Color.White

                    End If
                End If
            End If
        Next
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1 = New Timer
        Timer1.Interval = 20
        Timer1.Enabled = True
        ' notify() ' Check every 1 minute (60000 milliseconds)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnSort.Click
        DataGridView1.Sort(DataGridView1.Columns("PurchaseDate"), System.ComponentModel.ListSortDirection.Ascending)
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        LoadGroceryItemDataFromDatabase()
    End Sub

    Private Sub btnFilter_Click(sender As Object, e As EventArgs) Handles btnFilter.Click
        '  Dim selectedDateAdded As String = If(DateTimePicker1.Text IsNot Nothing, DateTimePicker1.Text.ToString(), "")
        Dim selectedGrocery As String = If(ComboBox2.SelectedItem IsNot Nothing, ComboBox2.SelectedItem.ToString(), "")

        Rinae.FilterGrocery(selectedGrocery) ', selectedDateAdded)
    End Sub
End Class