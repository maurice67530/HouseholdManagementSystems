Imports System.Data.OleDb

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try

            'Create an instance of Inventoryinformation 
            Dim inventory As New Inventory With {
                 .ItemName = TextBox1.Text,
                .Description = TextBox6.Text,
                .Quantity = TextBox3.Text,
                .Unit = TextBox2.Text,
                .Category = TextBox8.Text,
                .ReorderLevel = TextBox7.Text,
                .PricePerUnit = TextBox5.Text,
                .ExpiryDate = DateTimePicker1.Value,
                 .DateAdded = DateTimePicker2.Value}

            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()

                '' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                'Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                'Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Update the table name if necessary  
                Dim tableName As String = "Inventory"
                Dim query As String = "INSERT INTO Inventory(ItemName, Description, Quantity, Unit, Category, ReorderLevel, PricePerUnit, ExpiryDate, DateAdded) VALUES (@ItemName, @Description, @Quantity, @Unit, @Category, @ReorderLevel, @PricePerUnit, @ExpiryDate, @DateAdded)"
                Dim cmd As New OleDbCommand(query, conn)

                'params
                'cmd.Parameters.AddWithValue("@itemID", inventory.ItemID)
                'cmd.Parameters.AddWithValue("@ID", inventory.ID)
                cmd.Parameters.AddWithValue("@ItemName", inventory.ItemName)
                cmd.Parameters.AddWithValue("@Description", inventory.Description)
                cmd.Parameters.AddWithValue("@Quantity", inventory.Quantity)
                cmd.Parameters.AddWithValue("@Unit", inventory.Unit)
                cmd.Parameters.AddWithValue("@Category", inventory.Category)
                cmd.Parameters.AddWithValue("@ReorderLevel", inventory.ReorderLevel)
                cmd.Parameters.AddWithValue("@PricePerUnit", inventory.PricePerUnit)
                cmd.Parameters.AddWithValue("@expirydate", inventory.ExpiryDate)
                cmd.Parameters.AddWithValue("@DateAdded", inventory.DateAdded)

                Dim rowsAffected As Integer = cmd.ExecuteNonQuery
                MsgBox("inventory saved Successfuly!", vbInformation, "Update Confirmation")

            End Using


        Catch ex As OleDbException
            Debug.WriteLine($"Database error: {ex.Message}")
            MessageBox.Show($"Error saving inventory to database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Debug.WriteLine($"General error in btnSubmit_Click: {ex.Message}")
            MessageBox.Show($"An Unexpected error occured.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)



        End Try
    End Sub
End Class
