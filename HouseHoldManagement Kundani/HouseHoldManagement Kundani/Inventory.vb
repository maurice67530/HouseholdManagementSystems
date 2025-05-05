Imports System.Data.OleDb
Public Class Inventory
    Public Sub LoadInventorydataFromDatabase()
        Try
            '  Dim dataTable As DataTable = HouseHold.GetData("SELECT * FROM Expense")
            ' DataGridView1.DataSource = DataTable
            Debug.WriteLine("Populate Datagridview: Datagridview populated successfully.")
            Using conn As New OleDbConnection(Rinae.connectionString)
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
        Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
            Try
                Using conn As New OleDbConnection(Rinae.connectionString)

                    conn.Open()
                    Dim cmd As New OleDbCommand($"INSERT INTO Inventory ([ItemName], [Description], [Quantity], [Category], [ReorderLevel], [PricePerUnit], [DateAdded], [ExpiryDate], [Unit]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)", conn)

                    cmd.Parameters.Clear()
                    cmd.Parameters.AddWithValue("@ItemName", TextBox1.Text)
                    cmd.Parameters.AddWithValue("@Description", TextBox2.Text)
                    cmd.Parameters.AddWithValue("@Quantity", TextBox3.Text)
                    cmd.Parameters.AddWithValue("@Category", ComboBox1.SelectedItem)
                    cmd.Parameters.AddWithValue("@ReorderLevel", ComboBox2.SelectedItem)
                    cmd.Parameters.AddWithValue("@PricePerUnit", TextBox6.Text)
                    cmd.Parameters.AddWithValue("@DateAdded", DateTimePicker1.Value)
                    cmd.Parameters.AddWithValue("@ExpiryDate", DateTimePicker2.Value)
                    cmd.Parameters.AddWithValue("@Unit", ComboBox3.SelectedItem)

                    cmd.ExecuteNonQuery()
                    conn.Close()
                    LoadInventorydataFromDatabase()

                    MsgBox("Inventory Items Saved Successfuly!", vbInformation, "Inventory Saved")

                End Using
            Catch ex As OleDbException
                Debug.WriteLine($"Database error: {ex.Message}")
                Debug.Write($"Stack Trace: {ex.StackTrace}")
                MessageBox.Show("Error saving inventory to database: Please check the connectivity." & ex.Message & vbNewLine & ex.StackTrace, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Catch ex As Exception
                Debug.WriteLine($"General error: {ex.Message}")

                MessageBox.Show("Unexpected Error: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
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
                Dim Category As String = ComboBox1.Text
                Dim ReorderLevel As String = ComboBox3.Text
                Dim PricePerUnit As String = TextBox6.Text
                Dim DateAdded As String = DateTimePicker1.Value
                Dim ExpiryDate As String = DateTimePicker1.Value
                Dim Unit As String = ComboBox2.Text

                Using conn As New OleDbConnection(Rinae.connectionString)

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
                        Using conn As New OleDbConnection(Rinae.connectionString)
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

        Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
            Dim selectedCategory As String = If(ComboBox1.SelectedItem IsNot Nothing, ComboBox1.SelectedItem.ToString(), "")
            Dim selectedUnit As String = If(ComboBox2.SelectedItem IsNot Nothing, ComboBox2.SelectedItem.ToString(), "")

            Cruwza.FilterInventory(selectedCategory, selectedUnit)
        End Sub
        Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
            TextBox1.Text = ""
            TextBox3.Text = ""
            TextBox2.Text = ""
            TextBox6.Text = ""
            ComboBox3.Text = ""
            ComboBox1.SelectedItem = ""
            ComboBox2.SelectedItem = ""
        End Sub
        Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
            Try
                Debug.WriteLine("Selecting data in the GDV: Data selected")
                If DataGridView1.SelectedRows.Count > 0 Then
                    Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

                    ' Load the data from the selected row into UI controls  
                    TextBox1.Text = selectedRow.Cells("ItemName").Value.ToString()
                    TextBox2.Text = selectedRow.Cells("Description").Value.ToString()
                    TextBox3.Text = selectedRow.Cells("Quantity").Value.ToString()
                    ComboBox1.Text = selectedRow.Cells("Category").Value.ToString()
                    ComboBox3.Text = selectedRow.Cells("ReorderLevel").Value.ToString()
                    TextBox6.Text = selectedRow.Cells("PricePerUnit").Value.ToString()
                    ComboBox2.Text = selectedRow.Cells("Unit").Value.ToString()
                End If

                ' Enable/disable the buttons based on the selected person  
                Button1.Enabled = False

            Catch ex As Exception
                Debug.WriteLine("Data not selected: Error")
                Debug.Write($"Stack Trace: {ex.StackTrace}")
            End Try
        End Sub
        Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
            DataGridView1.Sort(DataGridView1.Columns("DateAdded"), System.ComponentModel.ListSortDirection.Ascending)
        End Sub
        Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
            LoadInventorydataFromDatabase()
        End Sub
        Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
            HighlightExpiredItems()
        End Sub

        Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
            ' Initially, mark expired groceries in red
            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not row.IsNewRow Then
                    Dim expiryDate As DateTime
                    If DateTime.TryParse(row.Cells("ExpiryDate").Value.ToString(), expiryDate) Then
                        If expiryDate < DateTime.Now.Date Then ' Only expired groceries (past expiry date)
                            row.DefaultCellStyle.ForeColor = Color.Red ' Mark expired groceries in red
                        End If
                    End If
                End If
            Next

            ' Initially, mark expired groceries in red
            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not row.IsNewRow Then
                    Dim expiryDate As DateTime
                    If DateTime.TryParse(row.Cells("ExpiryDate").Value.ToString(), expiryDate) Then

                        If expiryDate < DateTime.Now.Date Then ' Only expired groceries (past expiry date)
                            row.DefaultCellStyle.ForeColor = Color.Red ' Mark expired groceries in red

                        End If
                    End If
                End If
            Next
        End Sub
        'Modify the HighlightExpiredItems() method To handle NULL values properly
        Private Sub HighlightExpiredItems()
            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not row.IsNewRow Then
                    ' Check if the ExpirationDate column contains NULL before converting
                    If Not IsDBNull(row.Cells("ExpiryDate").Value) AndAlso Not String.IsNullOrEmpty(row.Cells("ExpiryDate").Value.ToString()) Then
                        Dim expDate As Date = CDate(row.Cells("ExpiryDate").Value)
                        If expDate < Date.Today Then
                            row.DefaultCellStyle.BackColor = Color.Red  ' Highlight expired items in red
                            row.DefaultCellStyle.ForeColor = Color.White ' Change text color for visibility
                        End If
                    End If
                End If
            Next
        End Sub

        Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
            Timer1.Stop()
            HighlightExpiredItems()
        End Sub

        Private Sub Inventory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            LoadInventorydataFromDatabase()

            ToolTip1.SetToolTip(Button1, "Save")
            ToolTip1.SetToolTip(Button2, "Edit")
            ToolTip1.SetToolTip(Button3, "Delete")
            ToolTip1.SetToolTip(Button4, "Clear")
            ToolTip1.SetToolTip(Button5, "Highlight")
            ToolTip1.SetToolTip(Button6, "Sort")
            ToolTip1.SetToolTip(Button8, "Refresh")
            ToolTip1.SetToolTip(Button7, "Filter")
            ToolTip1.SetToolTip(Button9, "Dashboard")
        End Sub

        Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
            'Dashboard.ShowDialog()
            Me.Close()
        End Sub

    End Class