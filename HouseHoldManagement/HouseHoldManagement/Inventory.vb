Public Class Inventory
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try

            'Create an instance of Inventoryinformation 
            Dim inventory As New Inventory1 With {
  .ItemName = TextBox1.Text,
                .Description = TextBox2.Text,
                .Quantity = TextBox3.Text,
                .Unit = TextBox7.Text,
                .Category = TextBox4.Text,
                .ReorderLevel = TextBox5.Text,
                .PricePerUnit = TextBox6.Text,
            .Totalcost = TextBox10.Text,
                .Total = TextBox9.Text,
                 .ExpiryDate = DateTimePicker1.Value,
            .datepurchased = DateTimePicker2.Value,
            .dateconsumed = DateTimePicker3.Value}

            Using conn As New OleDbConnection(callingmodule.connectionString)
                conn.Open()

                '' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                'Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                'Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Update the table name if necessary  
                Dim tableName As String = "INVENTORY"
                Dim query As String = "INSERT INTO INVENTORY (Name, Description, Quantity, Unit, Category, ReorderLevel, PricePerUnit, Totalcost, Total, expirydate, dateconsumed, datepurchased ) VALUES (@Name, @Discription, @Quantity, @Unit, @Category, @ReorderLevel, @PricePerUnit, @Totalcost, @Total, @expirydate, @dateconsumed, @datepurchased)"
                Dim cmd As New OleDbCommand(query, conn)

                'params
                'cmd.Parameters.AddWithValue("@itemID", inventory.ItemID)
                'cmd.Parameters.AddWithValue("@ID", inventory.ID)
                cmd.Parameters.AddWithValue("@Name", inventory.Name)
                cmd.Parameters.AddWithValue("@Description", inventory.Description)
                cmd.Parameters.AddWithValue("@Quantity", inventory.Quantity)
                cmd.Parameters.AddWithValue("@Unit", inventory.Unit)
                cmd.Parameters.AddWithValue("@Category", inventory.Category)
                cmd.Parameters.AddWithValue("@ReorderLevel", inventory.ReorderLevel)
                cmd.Parameters.AddWithValue("@PricePerUnit", inventory.PricePerUnit)
                cmd.Parameters.AddWithValue("@Totalcost", inventory.Totalcost)
                cmd.Parameters.AddWithValue("@Total", inventory.Total)
                cmd.Parameters.AddWithValue("@expirydate", inventory.expiryDate)
                cmd.Parameters.AddWithValue("@dateconsumed", inventory.dateconsumed)
                cmd.Parameters.AddWithValue("@datepurchased", inventory.datepurchased)

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