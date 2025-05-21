
Imports System.IO
Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class MealPlan
    ' Create a ToolTip object
    Private toolTip As New ToolTip()

    ' Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
    Public Property conn As New OleDbConnection(connectionString)
    ' Connection string using relative path to the database
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\KUNDANI TRADING\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb;Persist Security Info=False;"

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        If TextBox2.Text = "" Then
            MsgBox("Fill all the field", MsgBoxStyle.Exclamation)
            Exit Sub

        End If
        ' Get the selected meal
        Dim selectedMeal As String = If(Listbox1.SelectedItem IsNot Nothing, Listbox1.SelectedItem.ToString().Trim(), TextBox1.Text.Trim())

        ' Call the function to subtract ingredients if a meal is selected
        If Listbox1.SelectedItem IsNot Nothing Then
            SubtractIngredientsFromInventory(selectedMeal)
        End If ' 



        Try
            Dim Meals As New MealPlans() With {
                .MealID = TextBox4.Text,
                .NameOfMeal = TextBox1.Text,
                .MealPlan = Listbox1.Text,
                .StartDate = DateTimePicker1.Value,
                .EndDate = DateTimePicker2.Value,
                .TotalCalories = NumericUpDown1.Text,
                .Description = TextBox2.Text,
                 .Ingredients = ComboBox2.Text,
                 .FilePath = TextBox5.Text,
                .Frequency = ComboBox3.Text}


            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                conn.Open()

                'Update the table name if neccessary
                Dim tableName As String = "MealPlan"

                'create an OleDbCommand to insert the personalDetails data into the database
                Dim cmd As New OleDbCommand("INSERT INTO MealPlans (MealID, NameOfMeal, MealPlan, StartDate, EndDate, TotalCalories, Description, Ingredients, FilePath, Frequency) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", conn)
                cmd.Parameters.Clear()
                'cmd.Parameters.AddWithValue("@ID", ID)
                cmd.Parameters.AddWithValue("@?", TextBox4.Text)
                cmd.Parameters.AddWithValue("@?", TextBox1.Text)
                cmd.Parameters.AddWithValue("@?", Listbox1.Text)
                cmd.Parameters.AddWithValue("@?", DateTimePicker1.Value)
                cmd.Parameters.AddWithValue("@?", DateTimePicker2.Value)
                cmd.Parameters.AddWithValue("@?", NumericUpDown1.Text)
                cmd.Parameters.AddWithValue("@?", TextBox2.Text)
                cmd.Parameters.AddWithValue("@?", ComboBox2.Text)
                cmd.Parameters.AddWithValue("@?", TextBox5.Text)
                cmd.Parameters.AddWithValue("@?", ComboBox3.Text)

                'Execute the SQL Command to insert the data
                cmd.ExecuteNonQuery()


                ' Display a confirmation messageBox  indicating  the Inventory saved was successful
                MsgBox("Meal information saved successfully!", vbInformation, "save confirmation")

                'Next
            End Using

        Catch ex As OleDbException
            MessageBox.Show("Error Saving Meal to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("unexpected Error:" & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Check for inventory selection (No warnings after saving)

        If ComboBox2.SelectedItem IsNot Nothing Then

            Dim selectedItem As String = ComboBox2.SelectedItem.ToString()

            ' Database connection

            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)

                conn.Open()

                ' Fetch item details from Inventory1, including ExpiryDate

                Dim fetchcommand As New OleDbCommand("SELECT ItemName, Quantity, ExpiryDate FROM Inventory WHERE ItemName = ?", conn)

                fetchcommand.Parameters.AddWithValue("@ItemName", selectedItem)

                Using Readers As OleDbDataReader = fetchcommand.ExecuteReader()

                    If Readers.Read() Then

                        Dim ItemQuantity As Integer = Convert.ToInt32(Readers("Quantity")) ' Available total quantity

                        Dim ExpiryDate As Date

                        ' Check if ExpiryDate is valid

                        If Date.TryParse(Readers("ExpiryDate").ToString(), ExpiryDate) Then

                            ' If item is expired, update the label but **do not show a warning again**

                            If ExpiryDate < Date.Today Then

                                Label11.Text = "Available Stock:  (Expired)" & ItemQuantity.ToString()

                                Button1.Enabled = False

                            Else

                                ' If item is in stock and not expired

                                If ItemQuantity > 0 Then

                                    ' Update label without warnings

                                    Label11.Text = "Available Stock: " & ItemQuantity.ToString()

                                    Button1.Enabled = True

                                Else

                                    ' Update label for out-of-stock items

                                    Label11.Text = "Available Stock: 0"

                                    Button1.Enabled = False

                                End If

                            End If

                        End If

                    End If

                End Using

            End Using

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        'insure if row is selected in the Datagridview
        If DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("please select a record to Edit .", "No selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Try
            Dim MealID As String = TextBox4.Text()
            Dim MealPlan As String = Listbox1.Text()
            Dim StartDate As DateTime = DateTimePicker1.Value()
            Dim EndDate As DateTime = DateTimePicker2.Value()
            Dim TotalCalories As String = NumericUpDown1.Text()
            Dim Description As String = TextBox2.Text()
            Dim Ingredients As String = ComboBox2.Text()
            Dim Frequency As String = ComboBox3.Text()

            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                conn.Open()

                'Get the ID of the selected rows(assuming your table has a primary key named "ID" 
                Dim selectedRows As DataGridViewRow = DataGridView1.SelectedRows(0)


                Dim ID As Integer = Convert.ToInt32(selectedRows.Cells("ID").Value) 'change "ID" to your primary key column name

                'create an oleDbcommand to update the personnel data in the database
                Dim cmd As New OleDbCommand("UPDATE [MealPlans] SET [MealID]= ?, [MealPlan]= ?, [StartDate]= ?,  [EndDate] = ?, [TotalCalories] = ?, [Description] = ?, [Ingredients] = ?, [Frequency] =? WHERE [ID] = ?", conn)

                'set the paremeter values from the ui controls
                cmd.Parameters.AddWithValue("@MealID", TextBox4.Text)
                cmd.Parameters.AddWithValue("@MealPlan", Listbox1.Text)
                cmd.Parameters.AddWithValue("@StartDate", DateTimePicker1.Value)
                cmd.Parameters.AddWithValue("@EndDate", DateTimePicker2.Value)
                cmd.Parameters.AddWithValue("@TotalCalories", NumericUpDown1.Text)
                cmd.Parameters.AddWithValue("@Description", TextBox2.Text)
                cmd.Parameters.AddWithValue("@Ingredients", ComboBox2.Text)
                cmd.Parameters.AddWithValue("@Frequency", ComboBox3.Text)
                cmd.Parameters.AddWithValue("@ID", ID)
                ' Primary key for matching record

                'Execute the SQL command to update the data
                cmd.ExecuteNonQuery()

                ' Display a confirmation messageBox  indicating  the update was successful
                MsgBox("MealPlan updated successfully!", vbInformation, "update confirmation")

                'LoadMealPlansDatafromDatabase()
                'HouseHoldManagment_Module.ClearControls(Me)

            End Using
        Catch ex As OleDbException
            MessageBox.Show("Error updating MealPlan to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("Unexpected Error: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Check if there are any selected rows in the DataGridView for PersonalDetails  
        If DataGridView1.SelectedRows.Count > 0 Then
            ' Get the selected row  
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Assuming there is an "ID" column which is the primary key in the database  
            Dim MealPlansId As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  

            ' Confirm deletion  
            Dim confirmationResult As DialogResult = MessageBox.Show("Are you sure you want to delete this MealPlan?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Hand)
            If confirmationResult = DialogResult.Yes Then
                ' Proceed with deletion  
                Try
                    Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                        conn.Open()

                        ' Create the delete command  
                        Dim cmd As New OleDbCommand("DELETE FROM [MealPlans] WHERE [ID] = ?", conn)
                        cmd.Parameters.AddWithValue("@ID", MealPlansId) ' Primary key for matching record  

                        ' Execute the delete command  
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show("Mealplan deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ' Optionally refresh DataGridView or reload from database  

                        Else
                            MessageBox.Show("No MealPlan was deleted. Please check if the ID exists.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Hand)
                        End If
                    End Using

                Catch ex As Exception
                    MessageBox.Show($"An error occurred while deleting the mealplan: {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If

        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click


        Dim TotalCalories As String = If(ComboBox1.SelectedItem IsNot Nothing, ComboBox1.SelectedItem.ToString(), "")
        Dim StartDate As DateTime = If(DateTimePicker2.Text IsNot Nothing, DateTimePicker2.Value.ToString(), "")
        HouseHoldManagment_Module.FilterMealPlan(TotalCalories, StartDate)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                conn.Open()

                'Update the table name if neccessary
                Dim tablename As String = "MealPlans"

                'Create an OleDbCommand to select the data from the database
                Dim cmd As New OleDbCommand($"SELECT*FROM  {tablename}", conn)

                'create a DataAdapter and fill a DataTable
                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable()
                da.Fill(dt)

                'Bind the DataTable to the DataGridView
                DataGridView1.DataSource = dt

            End Using

        Catch ex As OleDbException
            'MessageBox.Show("$Error loading PersonalDetails data from database: {ex.message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show("$Error Loading Chore to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            'MessageBox.Show("$unexpected Error:  {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show("$unexpected Error:" & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        DataGridView1.Sort(DataGridView1.Columns("TotalCalories"), System.ComponentModel.ListSortDirection.Descending)

    End Sub

    Private Sub PopulateDataGridView()
        Try
            Debug.WriteLine("PopulatedDataGridView")

            ''Dim dataTable As DataTable = connect.GetData("SELECT *FROM Expense")
            'DataGridView1.DataSource = dataTable
            Debug.WriteLine("PopulateDataGridView: DataGridView populated successfully.")

            ' Clear existing rows  
            DataGridView1.Rows.Clear()
            ' Add each expense to the DataGridView  
            'For Each exp As Class1.Expenses In Expenses
            '    DataGridView1.Rows.Add(exp.ItemID, exp.Amount, exp.TagField, exp.Frequency, exp.PaymentMethod, exp.Currency, exp.ApprovalStatus, exp.DateOfExpense)
            'Next

        Catch ex As OleDbException
            Debug.WriteLine($"Database error in btnSubmit _Click: {ex.Message}")
            MessageBox.Show("error saving expense to database. please check the connection. ", "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)


        Catch ex As Exception
            Debug.WriteLine($"General error in btnsubmit_Click : {ex.Message}")
            Debug.WriteLine($"Error in populatingDataGridView: {ex.Message}")
            Debug.WriteLine($"stack Trace: {ex.StackTrace}")
        End Try
    End Sub
    Private Sub MealPlan_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'LoadffamilyMembersyDatafromDatabase()
        'CheckMealSelection()
        PopulateComboboxFromDatabase(ComboBox2)
        'DisplayMealPlansBasedOnSelection()
        LoadInventoryFromDatabase()
        LoadMealsFromDatabase()

        ' ComboBox1.Items.AddRange(New String() {"< 50", "500 - 1000", "> 1000"})
        DataGridView1.Columns.Clear() ' Clear existing columns if any  

        Dim connection As New OleDbConnection(connectionString)
        ' Check database connectivity  
        Try
            ' Create a new OleDbConnection object and open the connection  

            connection.Open()
            ' Display the connection status on a button with a green background  
            Button7.Text = "Connected"
            Button7.BackColor = Color.Green
            Button7.ForeColor = Color.White
        Catch ex As Exception
            ' Display the connection status on a button with a red background  
            Button7.Text = "Not Connected"
            Button7.BackColor = Color.Red
            Button7.ForeColor = Color.White

            ' Display an error message  
            MessageBox.Show("Error connecting to the database: " & ex.Message, "Database Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Close the database connection  
            connection.Close()
        End Try
        ' Initialize ToolTip properties (optional)
        ToolTip.AutoPopDelay = 5000
        ToolTip.InitialDelay = 500
        ToolTip.ReshowDelay = 200
        ToolTip.ShowAlways = True

        'Set tooltips for buttons
        ToolTip1.SetToolTip(Button1, "Save")
        ToolTip1.SetToolTip(Button2, "Update")
        ToolTip1.SetToolTip(Button3, "Delete")
        ToolTip1.SetToolTip(Button4, "Filter")
        ToolTip1.SetToolTip(Button5, "Refresh")
        ToolTip1.SetToolTip(Button8, "Sort")
        ToolTip1.SetToolTip(Button6, "Dash Board")

        ComboBox3.Items.Add("Day")
        ComboBox3.Items.Add("Week")
        ComboBox3.Items.Add("Month")


        PopulateComboboxFromDatabase(ComboBox2)
        LoadMealPlansDatafromDatabase()
    End Sub
    Private Sub DataGridView1_Selectionchanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Try
            If DataGridView1.SelectedRows.Count > 0 Then
                Dim SelectedRows As DataGridViewRow = DataGridView1.SelectedRows(0)

                ' lod the data from from the selected rows into UI contols
                TextBox4.Text = SelectedRows.Cells("MealID").Value.ToString
                TextBox2.Text = SelectedRows.Cells("Description").Value.ToString
                NumericUpDown1.Text = SelectedRows.Cells("TotalCalories").Value.ToString
                Listbox1.Text = SelectedRows.Cells("MealPlan").Value.ToString()
                DateTimePicker1.Value = SelectedRows.Cells("StartDate").Value.ToString()
                DateTimePicker2.Value = SelectedRows.Cells("EndDate").Value.ToString()
                ComboBox2.Text = SelectedRows.Cells("Ingredients").Value.ToString
                ComboBox3.Text = SelectedRows.Cells("Frequency").Value.ToString
            End If
            'code for submitting expense ...
        Catch ex As OleDbException
            Debug.WriteLine($"Database error in btnSubmit _Click: {ex.Message}")
            MessageBox.Show("error saving expense to database. please check the connection. ", "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)


        Catch ex As Exception
            Debug.WriteLine($"General error in btnsubmit_Click : {ex.Message}")
            'MessageBox.Show("An unexpected error occured. ", "error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged


        If ComboBox2.Text.ToString().StartsWith("Chicken Eggs") Then
            PictureBox1.Image = System.Drawing.Image.FromFile _
        ("C:\Users\KUNDANI TRADING\Pictures \OIP (3).jpg")
            TextBox5.Text = "C:\Users\KUNDANI TRADING\Pictures \OIP (3).jpg"

        End If

        If ComboBox2.Text.ToString().StartsWith("Beef") Then
            PictureBox1.Image = System.Drawing.Image.FromFile _
        ("C:\Users\KUNDANI TRADING\Pictures\OIP.jpg")
            TextBox5.Text = "C:\Users\KUNDANI TRADING\Pictures\OIP.jpg"

        End If
        If ComboBox2.Text.ToString().StartsWith("chicken feet") Then
            PictureBox1.Image = System.Drawing.Image.FromFile _
        ("C:\Users\KUNDANI TRADING\Pictures\th (3).jpg")
            TextBox5.Text = "C:\Users\KUNDANI TRADING\Pictures\th (3).jpg"

        End If


        If ComboBox2.Text.ToString().StartsWith("Samp") Then
            PictureBox1.Image = System.Drawing.Image.FromFile _
        ("C:\Users\KUNDANI TRADING\Pictures\download (2).jpg")
            TextBox5.Text = "C:\Users\KUNDANI TRADING\Pictures\download (2)jpg"

        End If
        If ComboBox2.Text.ToString().StartsWith("Beef Wors") Then
            PictureBox1.Image = System.Drawing.Image.FromFile _
        ("C:\Users\KUNDANI TRADING\Pictures\OIP (2).jpg")
            TextBox5.Text = "C:\Users\KUNDANI TRADING\Pictures\OIP (2).jpg"

        End If
        If ComboBox2.Text.ToString().StartsWith("CornFlakes") Then
            PictureBox1.Image = System.Drawing.Image.FromFile _
        ("C:\Users\KUNDANI TRADING\Pictures\download (3).jpg")
            TextBox5.Text = "C:\Users\KUNDANI TRADING\Pictures\download (3).jpg"
        End If


        If ComboBox2.Text.ToString().StartsWith("Ground Beef") Then
            PictureBox1.Image = System.Drawing.Image.FromFile _
        ("C:\Users\KUNDANI TRADING\Pictures\OIP (4).jpg")
            TextBox5.Text = "C:\Users\KUNDANI TRADING\Pictures\OIP (4).jpg"
        End If


        Dim selectedItem As String = ComboBox2.SelectedItem.ToString() ' Meal selected from ComboBox2
        DisplayMealPlanBasedOnSelection()

        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)

            conn.Open() ' Fetch item details from Inventory1, including ExpiryDate

            Dim fetchcommand As New OleDbCommand("SELECT ItemName, Quantity, Expirydate FROM Inventory WHERE ItemName = ?", conn)

            fetchcommand.Parameters.AddWithValue("@ItemName", selectedItem)

            Using Readers As OleDbDataReader = fetchcommand.ExecuteReader()

                If Readers.Read() Then

                    Dim ItemQuantity As Integer = Convert.ToInt32(Readers("Quantity")) ' Available total quantity

                    Dim expirydate As Date

                    ' Check if ExpiryDate is valid

                    If Date.TryParse(Readers("Expirydate").ToString(), expirydate) Then

                        ' If item is expired, show a message and prevent submission

                        If expirydate < Date.Today Then

                            MessageBox.Show("No meals can be made due to Expired Item.", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                            Label11.Text = "Available Stock: " & ItemQuantity.ToString() & " Expired"
                            Listbox1.Items.Clear()
                            Button1.Enabled = False

                            Exit Sub

                        End If

                    End If




                    ' If item is in stock and not expired

                    If ItemQuantity > 0 Then

                        ' Show warning if stock is below 6

                        If ItemQuantity < 6 Then

                            MessageBox.Show("Warning: Stock is below 6 for this item.", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                        End If

                        ' Display available stock

                        Label11.Text = "Available Stock: " & ItemQuantity.ToString()

                        Button1.Enabled = True
                        ' Add meal name to ListBox
                        Listbox1.Items.Add(selectedItem)
                        'MsgBox("Meal Suggestions have been prepared with current Greocery Items.", MsgBoxStyle.Information, "No Available Meals")


                    Else

                        ' Item is out of stock

                        MessageBox.Show("Item is out of stock.", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Stop)

                        Label11.Text = "Available Stock: 0"

                        Button1.Enabled = False

                        ' Add meal name to ListBox
                        Listbox1.Items.Clear()
                    End If
                Else
                    Listbox1.Items.Clear()

                End If


            End Using

        End Using

        ' Now, filter the items to add to the ListBox

        Dim listItems As New List(Of String)

        ' Open connection again to fetch all items for ListBox excluding expired ones

        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)

            conn.Open()

            Dim fetchcommand As New OleDbCommand("SELECT ItemName, Expirydate FROM Inventory", conn)

            Using Readers As OleDbDataReader = fetchcommand.ExecuteReader()

                While Readers.Read()

                    Dim ItemName As String = Readers("ItemName").ToString()

                    Dim itemExpiry As Date

                    If Date.TryParse(Readers("expirydate").ToString(), itemExpiry) Then

                        ' Only add non-expired items to the list

                        If itemExpiry >= Date.Today Then

                            listItems.Add(Name)

                        End If

                    End If

                End While

            End Using

        End Using

        ' Clear the ListBox and add the filtered items
        Listbox1.Items.Clear()
        Listbox1.Items.AddRange(listItems.ToArray())


        DisplayMealPlanBasedOnSelection()

        'Dim selectedMeal As String = ComboBox2.SelectedItem.ToString()
        'Dim usersWhoDislike As New List(Of String)

        'Using conn As New SqlConnection(connectionString)
        '    conn.Open()
        '    Dim query As String = "SELECT UserName FROM Users WHERE DislikedMeals LIKE @meal"
        '    Using cmd As New SqlCommand(query, conn)
        '        cmd.Parameters.AddWithValue("@meal", "%" & selectedMeal & "%")
        '        Using reader As SqlDataReader = cmd.ExecuteReader()
        '            While reader.Read()
        '                usersWhoDislike.Add(reader("UserName").ToString())
        '            End While
        '        End Using
        '    End Using
        'End Using

        'If usersWhoDislike.Count > 0 Then
        '    Dim message As String = "The following users do not eat " & selectedMeal & ":" & vbCrLf &
        '                                String.Join(", ", usersWhoDislike)
        '    MessageBox.Show(message, "Meal Restriction", MessageBoxButtons.OK, MessageBoxIcon.Information)
        'End If
    End Sub


    Public Sub PopulateComboboxFromDatabase(ByRef combobox As ComboBox)
        Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)

        Try
            Debug.WriteLine("Combobox Populated successfully ")

            ' 1. Open the database connection  
            conn.Open()

            ' 2. Retrieve the FirstName and LastName columns from the PersonalDetails table  
            Dim query As String = "SELECT ItemName FROM Inventory"
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            ' 3. Bind the retrieved data to the combobox  
            combobox.Items.Clear()
            While reader.Read()
                combobox.Items.Add($"{reader("ItemName")} ")
            End While
            ' 4. Close the database connection  
            reader.Close()
        Catch ex As Exception
            Debug.WriteLine($"Failed to PopulateComboboxFromDatabase:  {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("An unexpected error occurred during PopulatingComboboxFromDatabase.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            Debug.WriteLine("Exiting Populated Combobox")

            ' Close the database connection  
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub


    Private mealPlanData As DataTable
        Private currentRowIndex As Integer = 0
    Private Sub btnPrintMealPlan_Click_1(sender As Object, e As EventArgs) Handles btnPrintMealPlan.Click
        'Dim conn As New OleDbConnection
        'Dim adapter As New OleDb.OleDbDataAdapter(query, Conn)
        'Dim table As New DataTable()
        ''adapter.Fill(table)
        'DataGridView1.DataSource = table

        PrintDialog1.Document = PrintDocument1
        If PrintDialog1.ShowDialog() = DialogResult.OK Then
            LoadFilteredMealPlan() ' Load filtered data based on selected frequency
            If mealPlanData.Rows.Count > 0 Then
                PrintDocument1.Print()
            Else
                MessageBox.Show("No meal plans found for the selected period.", "Print Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End If

    End Sub
    ' Load filtered meal plan data based on frequency
    Private Sub LoadFilteredMealPlan()
        Using dbConnection As New OleDbConnection(HouseHoldManagment_Module.connectionString)
            Dim selectedFilter As String = ComboBox3.SelectedItem?.ToString()
            Dim query As String = "SELECT * FROM MealPlans WHERE Frequency = ? AND 1=1"
            Dim startDate As Date = Date.Today
            Dim endDate As Date = Date.Today

            If Not String.IsNullOrEmpty(selectedFilter) Then
                Select Case selectedFilter
                    Case "Day"
                        endDate = startDate
                    Case "Week"
                        endDate = startDate.AddDays(7)
                    Case "Month"
                        endDate = startDate.AddMonths(1)
                End Select
            End If

            ' Create command with parameters
            Dim cmd As New OleDbCommand(query, dbConnection)
            cmd.Parameters.AddWithValue("?", selectedFilter) ' Bind Frequency value

            Dim adapter As New OleDbDataAdapter(cmd)
            mealPlanData = New DataTable()
            adapter.Fill(mealPlanData)

            DataGridView1.DataSource = mealPlanData ' Display filtered data in DataGridView
        End Using
    End Sub
    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage


        Dim font As New Font("Arial", 12)
        Dim brush As New SolidBrush(Color.Black)
        Dim yPos As Integer = 100
        Dim leftMargin As Integer = e.MarginBounds.Left

        ' Print Title
        e.Graphics.DrawString("Meal Plan Report", New Font("Arial", 16, FontStyle.Bold), brush, leftMargin, yPos)
        yPos += 40

        ' Check if there is data to print
        If mealPlanData.Rows.Count = 0 Then
            e.Graphics.DrawString("No data available.", font, brush, leftMargin, yPos)
            Exit Sub
        End If

        ' Print filtered meal plan data
        For Each row As DataRow In mealPlanData.Rows
            e.Graphics.DrawString("MealPlan: " & row("MealPlan").ToString(), font, brush, leftMargin, yPos)
            yPos += 30
            e.Graphics.DrawString("Start Date: " & Convert.ToDateTime(row("StartDate")).ToShortDateString(), font, brush, leftMargin, yPos)
            yPos += 30
            e.Graphics.DrawString("End Date: " & Convert.ToDateTime(row("EndDate")).ToShortDateString(), font, brush, leftMargin, yPos)
            yPos += 30
            e.Graphics.DrawString("Total Calories: " & row("TotalCalories").ToString(), font, brush, leftMargin, yPos)
            yPos += 30
            e.Graphics.DrawString("Description: " & row("Description").ToString(), font, brush, leftMargin, yPos)
            yPos += 40
            e.Graphics.DrawString("Frequency: " & row("Frequency").ToString(), font, brush, leftMargin, yPos)
            yPos += 40
        Next
    End Sub
    ' Dictionary to store inventory items and their total quantities
    Dim InventoryItems As New Dictionary(Of String, Integer)()

    ' Function to load inventory data from the database
    Private Sub LoadInventoryFromDatabase()
        InventoryItems.Clear() ' Clear previous data

        Dim query As String = "SELECT ItemName, Quantity FROM Inventory"

        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(query, conn)
                Using reader As OleDbDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim itemName As String = reader("ItemName").ToString().Trim()
                        Dim Quantity As Integer = Convert.ToInt32(reader("Quantity"))


                        ' Add item to inventory dictionary
                        InventoryItems(itemName) = Quantity
                    End While
                End Using
            End Using
        End Using
    End Sub
    ' Dictionary to store meals and their ingredients
    Dim mealList As New Dictionary(Of String, List(Of String))()

    ' Function to load meal data from the database
    Private Sub LoadMealsFromDatabase()
        mealList.Clear()

        Dim query As String = "SELECT ItemName, Ingredients FROM Recipe"

        Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(query, conn)
                Using reader As OleDbDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim recipe As String = reader("ItemName").ToString().Trim()
                        Dim ingredients As String() = reader("Ingredients").ToString().Split(","c)

                        ' Add meal and its ingredients to the meal list
                        mealList(recipe) = ingredients.Select(Function(i) i.Trim()).ToList()
                    End While
                End Using
            End Using
        End Using
    End Sub
    ' Function to get meals based on available ingredients
    Public Function GetMealsBasedOnIngredients(selectedIngredients As String) As List(Of String)
        Dim mealPlan As New List(Of String)

        ' Loop through meals and check if they can be made
        For Each recipe As KeyValuePair(Of String, List(Of String)) In mealList
            If recipe.Value.Contains(selectedIngredients) Then
                Dim allIngredientsAvailable As Boolean = True

                ' Check if all ingredients exist in Inventory1 and have Total > 0
                For Each ingredients As String In recipe.Value
                    If Not InventoryItems.ContainsKey(ingredients) OrElse InventoryItems(ingredients) <= 0 Then
                        allIngredientsAvailable = False
                        Exit For ' Stop checking further
                    End If
                Next

                ' Add meal only if all ingredients are available
                If allIngredientsAvailable Then
                    mealPlan.Add(recipe.Key)
                End If
            End If
        Next

        ' If no meals can be made, add a message
        If mealPlan.Count = 0 Then
            mealPlan.Add("No meals can be made due to missing ingredients.")
        End If

        Return mealPlan
    End Function
    ' Function to display meal suggestions based on ComboBox2 selection
    Private Sub DisplayMealPlanBasedOnSelection()
        If ComboBox2.SelectedItem Is Nothing Then Exit Sub ' Ensure selection is made

        Dim selectedIngredient As String = ComboBox2.SelectedItem.ToString().Trim()
        Dim mealPlan As List(Of String) = GetMealsBasedOnIngredients(selectedIngredient)

        ' Clear ListBox before updating
        Listbox1.Items.Clear()

        ' Display meals or missing ingredients in ListBox1
        For Each meal As String In mealPlan
            Listbox1.Items.Add(meal)
        Next
    End Sub
    ' Save Selected Meal and Update Inventory
    Private Sub SaveMealSelection()
        If Listbox1.SelectedItem Is Nothing Then Exit Sub ' Ensure a meal is selected
        Dim selectedMeal As String = Listbox1.SelectedItem.ToString().Trim()

        ' Subtract all ingredients for the selected meal
        SubtractIngredientsFromInventory(selectedMeal)

        ' Refresh inventory data after update
        LoadInventoryFromDatabase()
    End Sub
    Private Sub Listbox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Listbox1.SelectedIndexChanged
        TextBox1.Enabled = False
        TextBox4.Enabled = False
    End Sub
    ' Check if ingredients are expired or out of stock
    Private Function CheckIngredientStatus(ingredient As String) As String
        ' Check if the ingredient is in stock
        If InventoryItems.ContainsKey(ingredient) Then
            If InventoryItems(ingredient) = -1 Then
                Return "expired"
            ElseIf InventoryItems(ingredient) <= 0 Then
                Return "out of stock"
            Else
                Return "available"
            End If
        Else
            Return "not found"
        End If
    End Function

    ' Check if a meal can be made based on the availability of ingredients
    Private Function AreIngredientsAvailableForMeal(mealName As String) As Boolean
        Dim mealIngredients As List(Of String) = mealList(mealName)
        For Each ingredient In mealIngredients
            Dim status As String = CheckIngredientStatus(ingredient)
            If status = "expired" OrElse status = "out of stock" Then
                Return False
            End If
        Next
        Return True
    End Function
    Private Sub SubtractIngredientsFromInventory(mealName As String)
        If mealList.ContainsKey(mealName) Then
            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                conn.Open()
                ' Loop through each ingredient of the selected meal
                For Each ingredient As String In mealList(mealName)
                    ' Skip ingredients that are expired
                    If InventoryItems.ContainsKey(ingredient) AndAlso InventoryItems(ingredient) > 0 AndAlso InventoryItems(ingredient) <> -1 Then
                        ' Subtract 1 from the inventory dictionary
                        InventoryItems(ingredient) -= 1
                        ' Prevent negative values
                        If InventoryItems(ingredient) < 0 Then InventoryItems(ingredient) = 0

                        ' Update the database with the new total
                        Dim query As String = "UPDATE Inventory SET Quantity = @Quantity WHERE ItemName = @ItemName"
                        Using cmd As New OleDbCommand(query, conn)
                            cmd.Parameters.AddWithValue("@Quantity", InventoryItems(ingredient))
                            cmd.Parameters.AddWithValue("@ItemName", ingredient)
                            cmd.ExecuteNonQuery()
                        End Using
                    End If
                Next
            End Using
        End If
    End Sub
    ' Get Meals Based on Available Ingredients (including expired and out-of-stock handling)
    Public Function GetMealsBasedOnIngredient(selectedIngredient As String) As List(Of String)
        Dim availableMeals As New List(Of String)() ' Meals that can be made
        Dim unavailableMeals As New List(Of String)() ' Meals that can't be made due to missing or expired ingredients
        Dim suggestions As New List(Of String)() ' List to suggest meals with available ingredients
        Dim feedbackMessage As String = ""

        ' Loop through meals and check if they can be made with the selected ingredient
        For Each meal As KeyValuePair(Of String, List(Of String)) In mealList
            If meal.Value.Contains(selectedIngredient) Then
                Dim missingIngredients As New List(Of String)()
                Dim expiredIngredients As New List(Of String)()

                ' Check each ingredient for availability or expiry
                For Each ingredient In meal.Value
                    Dim status As String = CheckIngredientStatus(ingredient)
                    If status = "out of stock" Then
                        missingIngredients.Add(ingredient)
                    ElseIf status = "expired" Then
                        expiredIngredients.Add(ingredient)
                    End If
                Next

                ' If there are missing or expired ingredients, mark the meal as unavailable
                If missingIngredients.Count > 0 OrElse expiredIngredients.Count > 0 Then
                    unavailableMeals.Add(meal.Key) ' Meal can't be made
                Else
                    availableMeals.Add(meal.Key) ' Meal can be made
                End If
            End If
        Next

        ' Handle unavailable meals (those with expired or out-of-stock ingredients)
        If unavailableMeals.Count > 0 Then
            ' Loop through each unavailable meal
            For Each meal In unavailableMeals
                Dim mealIngredients As List(Of String) = mealList(meal)
                Dim missingIngredients As New List(Of String)()
                Dim expiredIngredients As New List(Of String)()

                ' Classify missing and expired ingredients
                For Each ingredient In mealIngredients
                    Dim status As String = CheckIngredientStatus(ingredient)
                    If status = "out of stock" Then
                        missingIngredients.Add(ingredient) ' Out of stock
                    ElseIf status = "expired" Then ' Ingredient expired
                        expiredIngredients.Add(ingredient)
                    End If
                Next

                ' Prepare message to display for unavailable meal
                Dim message As String = $"You cannot make {meal} because the following ingredients is "

                ' Check if missing ingredients
                If missingIngredients.Count > 0 Then
                    message &= "out of stock: " & String.Join(", ", missingIngredients)
                End If

                ' Check if expired ingredients
                If expiredIngredients.Count > 0 Then
                    If missingIngredients.Count > 0 Then message &= " and "
                    message &= "expired: " & String.Join(", ", expiredIngredients)
                End If

                ' Add message to the feedbackMessage string
                feedbackMessage &= message & vbCrLf

                ' Suggest other meals with the selected ingredient but with available ingredients
                For Each otherMeal In mealList
                    If otherMeal.Value.Contains(selectedIngredient) AndAlso AreIngredientsAvailableForMeal(otherMeal.Key) Then
                        suggestions.Add(otherMeal.Key)
                    End If
                Next
            Next

            ' Show the feedback message in MessageBox
            MessageBox.Show(feedbackMessage, "Meal Cannot Be Made", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            ' Suggest meals with available ingredients
            If suggestions.Count > 0 Then
                Dim suggestionMessage As String = "You can still make the following meals with available ingredients: " & String.Join(", ", suggestions)
                MessageBox.Show(suggestionMessage, "Suggested Meals", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Else
            ' If meals can be made, show available meals
            feedbackMessage = "You can make the following meals: " & String.Join(", ", availableMeals)
            MessageBox.Show(feedbackMessage, "Meals Available", MessageBoxButtons.OK, MessageBoxIcon.Information)
            suggestions.AddRange(availableMeals)
        End If

        ' Clear the previous items from the ListBox
        Listbox1.Items.Clear()

        ' Display available meals or suggestions in ListBox
        For Each meal In suggestions
            Listbox1.Items.Add(meal)
        Next

        '' After meal is selected and made, subtract ingredients
        'If suggestions.Count > 0 Then
        '    SubtractIngredientsFromInventory(suggestions(0)) ' Subtract inventory after the first available meal is made
        'End If

        Return suggestions
    End Function
    Public Sub LoadMealplansDatafromDatabase()
        Try
            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                conn.Open()

                'Update the table name if neccessary
                Dim tablename As String = "Mealplans"

                'Create an OleDbCommand to select the data from the database
                Dim cmd As New OleDbCommand($"SELECT*FROM  {tablename}", conn)

                'create a DataAdapter and fill a DataTable
                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable()
                da.Fill(dt)

                'Bind the DataTable to the DataGridView
                DataGridView1.DataSource = dt

            End Using

        Catch ex As OleDbException
            'MessageBox.Show("$Error loading PersonalDetails data from database: {ex.message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show("$Error Loading Chore to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            'MessageBox.Show("$unexpected Error:  {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show("$unexpected Error:" & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub
    ' Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
    '    Dim selectedMeal As String = ComboBox2.SelectedItem.ToString()
    '    Dim usersWhoDislike As New List(Of String)

    '    Using conn As New SqlConnection(connectionString)
    '        conn.Open()
    '        Dim query As String = "SELECT UserName FROM Users WHERE DislikedMeals LIKE @meal"
    '        Using cmd As New SqlCommand(query, conn)
    '            cmd.Parameters.AddWithValue("@meal", "%" & selectedMeal & "%")
    '            Using reader As SqlDataReader = cmd.ExecuteReader()
    '                While reader.Read()
    '                    usersWhoDislike.Add(reader("UserName").ToString())
    '                End While
    '            End Using
    '        End Using
    '    End Using

    '    If usersWhoDislike.Count > 0 Then
    '        Dim message As String = "The following users do not eat " & selectedMeal & ":" & vbCrLf &
    '                                    String.Join(", ", usersWhoDislike)
    '        MessageBox.Show(message, "Meal Restriction", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '    End If
    'End Sub
End Class
