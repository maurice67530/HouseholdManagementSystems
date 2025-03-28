
Imports System.IO
Imports System.Windows.Forms
Imports System.Data.OleDb
Public Class GroceryItemvb

    Dim Grocery_Items As vb.GroceryItem

    Private Sub btnDashboard_Click(sender As System.Object, e As System.EventArgs) Handles btnDashboard.Click
        Me.Close()
    End Sub
    Private Sub btnSubmit_Click(sender As System.Object, e As System.EventArgs) Handles btnSubmit.Click
        'create an instance of groceryItem

        Try
            'Assign values from Textboxes to the  GroceryItem properties 
            Dim item As New GroceryItem With {
        .ItemName = TextBox1.Text,
        .Quantity = TextBox2.Text,
        .NutritionValue = TextBox3.Text,
        .Category = TextBox4.Text,
        .Price = TextBox5.Text,
        .Unit = TextBox6.Text,
        .ShelfLife = TextBox7.Text,
        .ExpiryDate = DateTimePicker1.Text,
        .ispurchased = TextBox8.Text}


            'add the Grocery to the list
            Using connect As New OleDbConnection(Module1.connectionString)
                Module1.conn.Open()

                Dim tablename As String = "GroceryItem"

                Dim cmd As New OleDbCommand("INSERT INTO [GroceryItem] ([ItemName],[Quantity],[NutritionValue],[Category],[Price],[Unit],[ShelfLife],[ExpiryDate],[Ispurchased]) VALUES (@ItemName ,@Quantity ,@NutritionValue, @Category, @Price, @Unit, @ShelfLife,@ExpiryDate,@Ispurchased)", Module1.conn)
                'set the parameter values from the UI controls
                'class declaration  


                'params
                cmd.Parameters.AddWithValue("@ItemName", item.ItemName)
                cmd.Parameters.AddWithValue("@Quantity", item.Quantity)
                cmd.Parameters.AddWithValue("NutritionValue@", item.NutritionValue)
                cmd.Parameters.AddWithValue("@Category", item.Category)
                cmd.Parameters.AddWithValue("@Price", item.Price)
                cmd.Parameters.AddWithValue("@Unit", item.Unit)
                cmd.Parameters.AddWithValue("@ShelfLife", item.ShelfLife)
                cmd.Parameters.AddWithValue("@ExpiryDate", item.ExpiryDate)
                cmd.Parameters.AddWithValue("@Ispurchased", item.ispurchased)

                'cmd.ExecuteNonQuery()


            End Using
            'Display a confirmation message 
            MsgBox("Grocery Item Added!" & vbCrLf &
        "ItemName:" & item.ItemName & vbCrLf &
        "Quantity:" & item.Quantity & vbCrLf &
        "NutritionValue:" & item.NutritionValue & vbCrLf &
        "Category:" & item.Category & vbCrLf &
        "Price:" & item.Price & vbCrLf &
        "Unit:" & item.Unit & vbCrLf &
        "ShelLife:" & item.ShelfLife & vbCrLf &
        "ExpiryDate:" & item.ExpiryDate & vbCrLf &
        "ispurchased:" & item.ispurchased, vbInformation, "Item confirmation")



        Catch ex As OleDbException
            Debug.WriteLine($"Database error: {ex.Message}")
            'MessageBox.Show("Error saving test to database. Please check the connection.", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine($"stack Trace:{ex.StackTrace}")
        Catch ex As Exception
            Debug.WriteLine($"General error: {ex.Message}")
            'MessageBox.Show("An unexpected error occurred.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Module1.conn.Close()
        Finally

        End Try
        'populatedDataGridview()0
    End Sub

    'Private Sub populatedDataGridview()
    '    'clear Existing rows
    '    DataGridView1.Rows.Clear()

    '    'Add each expense to the DataGridView
    '    For Each per As GroceryItem In groceryItem
    '        DataGridView1.Rows.Add(per.ItemName, per.Quantity, per.NutritionValue, per.Category, per.Price, per.Unit, per.ShelfLife.ToString, per.ExpiryDate, per.ispurchased)

    '    Next
    'End Sub

    Public Sub LoadGroceryItemsFromDatabase()
        Try
            Debug.WriteLine("populated connected succesfully")
            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()

                Dim tableName As String = "GroceryItem"
                Dim cmd As New OleDbCommand($"SELECT * FROM {tableName}", conn)

                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)

                DataGridView1.DataSource = dt

            End Using

        Catch ex As Exception

            Debug.WriteLine("populated fail to initialize succesfully")
            MessageBox.Show($"Error Loading  Expenses data from database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        'Load the data from  the selected row into UI controls 
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
            TextBox1.Text = selectedRow.Cells("ItemName").Value.ToString()
            TextBox2.Text = selectedRow.Cells("Quantity").Value.ToString()
            'TextBox3.Text = selectedRow.Cells("NutriationValue").Value.ToString()
            TextBox4.Text = selectedRow.Cells("Category").Value.ToString()
            TextBox5.Text = selectedRow.Cells("Price").Value.ToString()
            TextBox6.Text = selectedRow.Cells("Unit").Value.ToString()
            'TextBox7.Text = selectedRow.Cells("ShelLife").Value.ToString()
            'DateTimePicker1.Text = selectedRow.Cells("ExpiryDate").Value.ToString()
            TextBox8.Text = selectedRow.Cells("ispurchased").Value.ToString()
        End If
    End Sub



    Private Sub GroceryItem_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        LoadGroceryItemsFromDatabase()
    End Sub
End Class




