
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
        LoadGroceryItemDataFromDatabase()
        ToolTip1.SetToolTip(btnSubmit, "Submit")

    End Sub
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged


        'Debug.WriteLine("Entering DataGridView1_SelectionChange")


        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Load the data from the selected row into UI controls
            TextBox1.Text = selectedRow.Cells("ItemName").Value.ToString()
            TextBox2.Text = selectedRow.Cells("Quantity").Value.ToString()
            ComboBox2.SelectedItem = selectedRow.Cells("Unit").Value.ToString()
            TextBox6.Text = selectedRow.Cells("Category").Value.ToString(
                TextBox5.Text = selectedRow.Cells("Price").Value.ToString())
            ComboBox1.SelectedItem = selectedRow.Cells("Ispurchased").Value.ToString()
            DateTimePicker1.Value = selectedRow.Cells("ExpiryDate").Value.ToString


            'Debug.WriteLine("Existing DataGridView1_SelectionChange")
            ' Enable/disable the buttons based on the selected person  
            btnSubmit.Enabled = False
            LoadGroceryItemDataFromDatabase()


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
        ToolTip1.SetToolTip(Button5, "Update")
        ToolTip1.SetToolTip(btnSubmit, "Submit")
        ToolTip1.SetToolTip(Button4, "Delete")
        ToolTip1.SetToolTip(Button3, "DashBoard")
        HighlightExpiredItems()
        'AlertExpiringGroceries()
        Timer1.Interval = 6000
        Timer1.Start()


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            ' Get the selected row  
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Assuming there is an "ID" column which is the primary key in the database  
            Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  

            ' Confirm deletion  
            Dim confirmationResult As DialogResult = MessageBox.Show("Are you sure you want to delete this GroceryItem?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If confirmationResult = DialogResult.Yes Then
                ' Proceed with deletion  
                Try
                    Using conn As New OleDbConnection(Rotondwa.connectionString)
                        conn.Open()

                        ' Create the delete command  
                        Dim cmd As New OleDbCommand("DELETE FROM [GroceryItem] WHERE [ID] = ?", conn)
                        cmd.Parameters.AddWithValue("@ID", ID) ' Primary key for matching record  

                        ' Execute the delete command  
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show("Grocery deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ' Optionally refresh DataGridView or reload from database  
                            ' PopulateDataGridView()
                        Else
                            MessageBox.Show("No Grocery was deleted. Please check if the ID exists.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End Using

                    LoadGroceryItemDataFromDatabase()

                Catch ex As Exception
                    Debug.WriteLine("Existing btnDelete")
                    Debug.WriteLine($"error in btnDelete_Click: {ex.Message}")
                    Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
                    MessageBox.Show($"An error occurred while deleting the Grocery: {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            MessageBox.Show("Please select an Mealplan to delete.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        ToolTip1.SetToolTip(Button4, "Delete")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' Ensure a row is selected in the DataGridView  
        If DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Debug.WriteLine("Entering btnUpdate")
            Using conn As New OleDbConnection(Rotondwa.connectionString)
                conn.Open()

                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Create an OleDbCommand to update the personnel data in the database  
                ' Create an OleDbCommand to update the personnel data in the database  
                Dim cmd As New OleDbCommand("UPDATE GroceryItem SET [ItemName] = ?, [Quantity] = ?, [Unit] = ?, [Category] = ?, [Price] = ?, [IsPurchased] = ?, [ExpiryDate] = ? WHERE ID = ?", conn)
                ' Set the parameter values from the UI controls 
                cmd.Parameters.AddWithValue("@?", TextBox1.Text)
                cmd.Parameters.AddWithValue("@?", TextBox2.Text)
                cmd.Parameters.AddWithValue("@?", ComboBox1.SelectedItem.ToString)
                cmd.Parameters.AddWithValue("@?", ComboBox2.SelectedItem.ToString)
                cmd.Parameters.AddWithValue("@?", TextBox5.Text)
                cmd.Parameters.AddWithValue("@?", TextBox6.Text)
                cmd.Parameters.AddWithValue("@?", DateTimePicker1.Value.ToLongDateString)
                cmd.ExecuteNonQuery()


                ' Execute the SQL command to update the data  
                cmd.ExecuteNonQuery()

                ' Displaya a message box indicate the upate was successful
                MsgBox("MealPlan Information Updated!", vbInformation, "Update Confirmation")

                ' Clear the controls and refresh the DataGridView  

                LoadGroceryItemDataFromDatabase() ' Reload data to reflect updates

                'New update code to be expoanded....inputbox will do

                Dim updateValues As New Dictionary(Of String, Object)
                updateValues.Add("Column1", "NewValue1")
                updateValues.Add("Column2", "NewValue2")
                'Rotondwa.updateData("YourTableName", updateValues, "ID = 1")

                'Rotondwa.ClearControls(Me)
                LoadGroceryItemDataFromDatabase() ' Reload data to reflect updates
                MsgBox("MealPlan  Updated Successfuly!", vbInformation, "Update Confirmation")

            End Using

        Catch ex As OleDbException
            Debug.WriteLine($" error in btnUpdate_Click: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error updating MealPlan in database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        ToolTip1.SetToolTip(Button5, "Update")
    End Sub
    Private Sub HighlightExpiredItems()
        ' Loop through each row in DataGridView
        For Each row As DataGridViewRow In DataGridView1.Rows
            ' Check if the expiry date is less than current date
            If row.Cells("ExpiryDate").Value IsNot Nothing Then
                Dim expiryDate As DateTime
                If DateTime.TryParse(row.Cells("ExpiryDate").Value.ToString(), expiryDate) Then
                    If expiryDate < DateTime.Now Then
                        ' Change the background color to red for expired items
                        row.DefaultCellStyle.BackColor = Color.Red
                    Else
                        ' Reset the color for non-expired items
                        row.DefaultCellStyle.BackColor = Color.White
                    End If
                End If
            End If
        Next
    End Sub
    Private Sub AlertExpiringGroceries()
        ' Variable to track if any item is expiring in the next 5 days
        Dim expiringItemsFound As Boolean = False

        ' Loop through each row in DataGridView
        For Each row As DataGridViewRow In DataGridView1.Rows
            ' Check if the expiry date is valid
            If row.Cells("ExpiryDate").Value IsNot Nothing Then
                Dim expiryDate As DateTime
                If DateTime.TryParse(row.Cells("ExpiryDate").Value.ToString(), expiryDate) Then
                    ' Check if the expiry date is within the next 5 days
                    If expiryDate > DateTime.Now AndAlso expiryDate <= DateTime.Now.AddDays(5) Then
                        ' Highlight the row if it is expiring in the next 5 days
                        row.DefaultCellStyle.BackColor = Color.Yellow

                        ' Set flag to indicate that expiring items were found
                        expiringItemsFound = True
                    Else
                        '' Reset the color for items not expiring soon
                        'row.DefaultCellStyle.BackColor = Color.White
                    End If
                End If
            End If
        Next

        ' Display alert if any item is expiring in the next 5 days
        If expiringItemsFound Then
            MsgBox("Some groceries are expiring in the next 5 days! Please check them.", MsgBoxStyle.Information, "Expiry Alert")
        Else
            MsgBox("No groceries are expiring in the next 5 days.", MsgBoxStyle.Information, "Expiry Alert")
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        AlertExpiringGroceries()
        AlertExpiredItems()
        Timer1.Interval = 6000
        Timer1.Stop()

    End Sub
    Private Sub AlertExpiredItems()
        ' Variable to track if any expired items are found
        Dim expiredItemsList As New List(Of String)

        ' Loop through each row in DataGridView
        For Each row As DataGridViewRow In DataGridView1.Rows
            ' Check if the expiry date is valid
            If row.Cells("ExpiryDate").Value IsNot Nothing Then
                Dim expiryDate As DateTime
                If DateTime.TryParse(row.Cells("ExpiryDate").Value.ToString(), expiryDate) Then
                    ' Check if the expiry date has passed
                    If expiryDate < DateTime.Now Then
                        ' Add the item name and expiry date to the list of expired items
                        Dim itemName As String = row.Cells("ItemName").Value.ToString() ' Adjust column name as needed
                        expiredItemsList.Add(itemName & " - Expired on: " & expiryDate.ToString("MM/dd/yyyy"))
                    End If
                End If
            End If
        Next

        ' If there are expired items, show them in a warning message box
        If expiredItemsList.Count > 0 Then
            Dim itemsMessage As String = "The following groceries have expired:" & vbCrLf
            For Each item As String In expiredItemsList
                itemsMessage &= item & vbCrLf
            Next
            MsgBox(itemsMessage, MsgBoxStyle.Exclamation, "Expired Items Warning")
        Else
            MsgBox("No groceries have expired.", MsgBoxStyle.Information, "Expiry Check")
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

        'If DataGridView1.SelectedRows.Count > 0 Then
        '    Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

        '    ' Load the data from the selected row into UI controls
        '    TextBox1.Text = selectedRow.Cells("ItemName").Value.ToString()
        '    TextBox2.Text = selectedRow.Cells("Quantity").Value.ToString()
        '    ComboBox1.SelectedItem = selectedRow.Cells("Unit").Value.ToString()
        '    TextBox6.Text = selectedRow.Cells("Category").Value.ToString(
        '        TextBox5.Text = selectedRow.Cells("Price").Value.ToString())
        '    TextBox6.Text = selectedRow.Cells("Ispurchased").Value.ToString()
        '    DateTimePicker1.Value = selectedRow.Cells("ExpiryDate").Value.ToString


        '    'Debug.WriteLine("Existing DataGridView1_SelectionChange")
        '    ' Enable/disable the buttons based on the selected person  
        '    btnSubmit.Enabled = False
        '    LoadGroceryItemDataFromDatabase()


        'End If
    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ToolTip1.SetToolTip(Button3, "DashBoard")
    End Sub
End Class