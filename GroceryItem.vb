
Imports System.IO
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Drawing

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

            'Dim selectedItem As String = ComboBox1.SelectedItem.ToString() 'Meal selected from combobox2

            'Using conn As New OleDbConnection(Module1.connectionString)
            '    conn.Open()

            '    'Fetch the item and its details from the inventory table

            '    Dim fetchCommand As New OleDbCommand("SELECT ItemName,Unit,PriceperUnit FROM Inventory WHERE ItemName=?", conn)
            '    fetchCommand.Parameters.AddWithValue("@ItemName", selectedItem)

            '    Using Readers As OleDbDataReader = fetchCommand.ExecuteReader()
            '        If Readers.Read() Then
            '            Dim ItemQuantity As Integer = Convert.ToInt32(Readers("Unit")) 'Get available total quantity of the item
            '            'If the item is in stock
            '            If ItemQuantity > 0 Then
            '                'Update the inventory by reducing quantity by 1(the  meal uses one unit of the item

            '                Dim UpdateCommand As New OleDbCommand("UPDATE Inventory SET Unit=Unit-1 WHERE ItemName=?", conn)
            '                UpdateCommand.Parameters.AddWithValue("@Itemname", selectedItem)

            '                UpdateCommand.ExecuteNonQuery()

            '                'Display confirmation message 
            '                MessageBox.Show("Item added to MealPlan.Inventory updated.")

            '                'Dispaly updated stock Status
            '                TextBox10.Text = "Available Stock:" & (ItemQuantity - 1).ToString
            '            Else

            '                'Item is out of stock
            '                MessageBox.Show("Item is out of stock.")
            '                TextBox10.Text = "Available Stock:0"
            '            End If
            '        End If
            '    End Using
            'End Using

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

    'Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
    '    If DataGridView1.Columns(e.ColumnIndex).Name = "ExpiryDate" Then
    '        Dim expiryDate As Date = CDate(e.Value)
    '        If expiryDate < Date.Now Then
    '            e.CellStyle.BackColor = Color.Red
    '            e.CellStyle.ForeColor = Color.White
    '        End If
    '    End If
    'End Sub

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
                HighlightExpiredItemss()
            End Using

        Catch ex As Exception

            Debug.WriteLine("populated fail to initialize succesfully")
            MessageBox.Show($"Error Loading  Expenses data from database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub HighlightExpiredItemss()
        Dim expiredItems As New List(Of String)
        For Each row As DataGridViewRow In DataGridView1.Rows
            If Not row.IsNewRow Then
                Dim expDate As DateTime = (row.Cells("ExpiryDate").Value)
                If expDate < Date.Today Then
                    row.DefaultCellStyle.BackColor = Color.Maroon ' Highlight expired items in red
                    row.DefaultCellStyle.ForeColor = Color.White ' Change text color for visibility
                    expiredItems.Add(row.Cells("ItemName").Value.ToString()) ' Store expired item name
                End If
            End If
        Next

        ' Show message box if there are expired items
        If expiredItems.Count > 0 Then
            MessageBox.Show("The following Highlighted Items have expired:" & vbCrLf & String.Join(vbCrLf, expiredItems),
                                    "Expired Items Alert", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        End If
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
    Private Sub PopulateListboxFromdatabase()


        Dim conn As New OleDbConnection(Module1.connectionString)
        Try
            Debug.WriteLine("Listbox populated from database")
            conn.Open()

            'Retrieve the status and title columns from the chore tabel
            Dim query As String = "SELECT ItemName,ShelfLife From GroceryItem "
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader

            'bind the retrieved data to the combobox 
            ListBox1.Items.Clear()
            While reader.Read()
                ListBox1.Items.Add($"{reader("ItemName")}")
            End While

            'Close the database
            reader.Close()

        Catch ex As Exception
            Debug.WriteLine("Failed to populate listbox")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            'handle any exeptions that may occur
            MessageBox.Show($"Error:{ex.Message}")

        Finally

            'close the database connection 
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
        Debug.WriteLine("Done with populating listbox from database ")
    End Sub
    'Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
    '    Dim selectedItem As String = ComboBox1.SelectedItem.ToString() ' Meal selected from ComboBox1

    '    Using conn As New OleDbConnection(Module1.connectionString)
    '        conn.Open()

    '        ' Fetch the item and its details from Inventory1 table.
    '        Dim fetchcommand As New OleDbCommand("SELECT ItemName, Unit,PriceperUnit FROM Inventory  WHERE ItemName = ?", conn)
    '        fetchcommand.Parameters.AddWithValue("@ItemName", selectedItem)

    '        Using Readers As OleDbDataReader = fetchcommand.ExecuteReader()
    '            If Readers.Read() Then
    '                Dim ItemQuantity As Integer = Convert.ToInt32(Readers("Unit")) ' Get available total quantity of the item
    '                Dim ItemTotalCost As Integer = Convert.ToInt32(Readers("PriceperUnit")) ' Get total cost for the item

    '                ' If the item is in stock
    '                If ItemQuantity > 0 Then
    '                    ' Display the calories or total cost for the selected item
    '                    TextBox9.Text = "PriceperUnit: " & ItemTotalCost.ToString()

    '                    ' If stock is below 6, show a warning
    '                    If ItemQuantity < 6 Then
    '                        MessageBox.Show("Warning: Stock is below 6 for this item.")
    '                    End If

    '                    ' Display the current available stock
    '                    TextBox6.Text = "Available Stock: " & ItemQuantity.ToString()
    '                Else
    '                    ' Item is out of stock
    '                    MessageBox.Show("Item is out of stock.")
    '                    TextBox9.Text = "Available Stock: 0"
    '                End If
    '            End If
    '        End Using
    '    End Using
    'End Sub

    Private Sub GroceryItem_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        PopulateListboxFromdatabase()
        LoadGroceryItemsFromDatabase()

        ' Set tooltips For buttons
        ToolTip1.SetToolTip(btnDashboard, "Dashboard")
        ToolTip1.SetToolTip(btnSubmit, "Submit")
        ToolTip1.SetToolTip(btnClear, "Clear")
        ToolTip1.SetToolTip(Button1, "Edit")
        ToolTip1.SetToolTip(Button2, "Delete")
        ToolTip1.SetToolTip(Button5, "Sort")
        ToolTip1.SetToolTip(Button6, "Filter")
        ToolTip1.SetToolTip(Button7, "Highlight")
        ToolTip1.SetToolTip(Button8, "Calculate")

        '' clear Any existing items in combobox1 In ComboBox1

        'ComboBox2.Items.Clear()
        'Using conn As New OleDbConnection(Module1.connectionString)
        '    conn.Open()
        '    'Query to fetch all ItemName Values from Inventory
        '    Dim fetchcommand As New OleDbCommand("SELECT ItemName FROM Inventory ", conn)

        '    Using Readers As OleDbDataReader = fetchcommand.ExecuteReader()
        '        'Add ItemName Values to Combobox1
        '        While Readers.Read()
        '            ComboBox2.Items.Add(Readers("ItemName").ToString())


        'End While
        '    End Using
        'End Using
    End Sub
End Class




