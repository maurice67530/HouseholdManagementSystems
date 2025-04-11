
Imports System.IO
Imports System.Data.OleDb
Public Class groceryForm
    Dim conn As New OleDbConnection(Rotondwa.connectionString)
    Private Grocery As New List(Of GroceryClass)

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click


        'create an instance of GroceryItem
        Dim Item As New GroceryClass

        'assign values from textboxes to the GroceryItems properties
        Item.ItemName = TextBox1.Text
        Item.Quantity = TextBox2.Text
        Item.Ispurchased = TextBox6.Text
        Item.Category = ComboBox2.Text
        Item.Price = TextBox5.Text
        Item.Unit = ComboBox1.Text

        ' Add the GroceryItem to the list  
        Grocery.Add(Item)



        Debug.WriteLine("Entering btnSubmit")

        Using conn As New OleDbConnection(Rotondwa.connectionString)
            conn.Open()
            Dim cmd As New OleDbCommand("INSERT INTO GroceryItem ([ItemName], [Quantity], [Unit], [Category], [Price], [Ispurchased], [ExpiryDate]) VALUES (?,?,?,?,?,?,?)", conn)
            cmd.Parameters.AddWithValue("@?", TextBox1.Text)
            cmd.Parameters.AddWithValue("@?", TextBox2.Text)
            cmd.Parameters.AddWithValue("@?", ComboBox1.SelectedItem.ToString)
            cmd.Parameters.AddWithValue("@?", ComboBox2.SelectedItem.ToString)
            cmd.Parameters.AddWithValue("@?", TextBox5.Text)
            cmd.Parameters.AddWithValue("@?", TextBox6.Text)
            cmd.Parameters.AddWithValue("@?", DateTimePicker1.Value.ToLongDateString)
            cmd.ExecuteNonQuery()

        End Using

        'display a confirmation message
        MsgBox("Grocery Items Added!" & vbCrLf & "ItemName:" & Item.ItemName & vbCrLf & "Quantity:" & Item.Quantity & vbCrLf & "Ispurchased:" & Item.Ispurchased & vbCrLf & "Category:" & Item.Category & vbCrLf & "Price:" & Item.Price.ToString & vbCrLf & "Unit:" & Item.Unit & vbCrLf & "ItemID:" & Item.ItemID, vbInformation, "Item Confirmation")
        'LoadGroceryItemDataFromDatabase()


    End Sub
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged


        'Debug.WriteLine("Entering DataGridView1_SelectionChange")


        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Load the data from the selected row into UI controls
            TextBox1.Text = selectedRow.Cells("ItemName").Value.ToString()
            TextBox2.Text = selectedRow.Cells("Quantity").Value.ToString()
            ComboBox1.SelectedItem = selectedRow.Cells("Unit").Value.ToString()
            TextBox5.Text = selectedRow.Cells("Category").Value.ToString(
                TextBox5.Text = selectedRow.Cells("Price").Value.ToString())
            TextBox6.Text = selectedRow.Cells("Ispurchased").Value.ToString()
            DateTimePicker1.Value = selectedRow.Cells("ExpiryDate").Value.ToString


            'Debug.WriteLine("Existing DataGridView1_SelectionChange")
            ' Enable/disable the buttons based on the selected person  
            btnSubmit.Enabled = False



        End If
    End Sub
    Public Sub LoadGroceryItemDataFromDatabase()





        Try
            Debug.WriteLine("Loading successfully")
            Using conn As New OleDbConnection(Rotondwa.connectionString)

                conn.Open()

                ' Update the table name if necessary
                Dim tableName As String = "GroceryItem"

                ' Creatan OleDbCommand to select the data from the database
                Dim cmd As New OleDbCommand($"SELECT*FROM {tableName}", conn)

                'Create a DataAdapter and fill a Data Table 
                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable()
                da.Fill(dt)

                'Bind the data Table to the datagridview
                DataGridView1.DataSource = dt
            End Using


        Catch ex As OleDbException
            Debug.WriteLine("Failed To Load")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"UnExpected error in loading Grocery   from database.{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub groceryForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadGroceryItemDataFromDatabase()

    End Sub
End Class