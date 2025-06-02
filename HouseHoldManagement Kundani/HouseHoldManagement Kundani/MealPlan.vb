
Imports System.IO
Imports System.Data.OleDb

Public Class MealPlan


    Public Property conn As New OleDbConnection(connectionString)
    'Public Const connectionString As String = " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Nedzamba\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"
    Dim eventType As String = "MealPlans"
    Dim cmd As OleDbCommand
    Dim da As OleDbDataAdapter
    Dim dt As DataTable

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        Try
            Debug.WriteLine("Entering btnEdit_Click")
            Using conn As New OleDbConnection(connectionString)
                conn.Open()

                Dim tablename As String = "MealPlans"
                Using cmd As New OleDbCommand("INSERT INTO MealPlans ([StartDate], [EndDate], [Meals], [MealName], [Items], [TotalCalories], [Description], [FilePath], [Calories], [Frequency]) VALUES (@StartDate, @EndDate, @Meals, @MealName, @Items, @TotalCalories, @Description, @FilePath, @Calories, @Frequency)", conn)

                    cmd.Parameters.AddWithValue("@StartDate", DateTimePicker1.Text)
                    cmd.Parameters.AddWithValue("@EndDate", DateTimePicker2.Text)
                    cmd.Parameters.AddWithValue("@Meals", lstMealSuggestions.SelectedItem.ToString)
                    cmd.Parameters.AddWithValue("@MealName", TextBox4.Text)
                    cmd.Parameters.AddWithValue("@Items", ComboBox3.SelectedItem.ToString)
                    cmd.Parameters.AddWithValue("@TotalCalories", NumericUpDown1.Value)
                    cmd.Parameters.AddWithValue("@Description", TextBox2.Text)
                    cmd.Parameters.AddWithValue("@FilePath", TextBox3.Text)
                    cmd.Parameters.AddWithValue("@Calories", ComboBox1.SelectedItem.ToString)
                    cmd.Parameters.AddWithValue("@Frequency", ComboBox2.SelectedItem.ToString)
                    cmd.ExecuteNonQuery()
                End Using
                MessageBox.Show("MealPlan Updated successfully")

            End Using

            Debug.WriteLine("The data has been edited successfully")
        Catch ex As OleDbException
            Debug.WriteLine($"Database saving in btnEdit_Click: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            ' MessageBox.Show("error saving expense to database. please check the connection.", "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine($"General error in btnEdit_Click: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
        End Try

        Debug.WriteLine("Existing btnSave_Click")

    End Sub

    Private Sub MealPlan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.AddRange(New String() {"<500", "500-1000", ">1000"})
        ComboBox2.Items.AddRange(New String() {"Daily", "Weekly", "Monthly"})
        CheckDatabaseConnection(StatusLabel)
        lstMealSuggestions.Items.AddRange(New String() {"Noodles", "Chicken Curry", "Kota"})
        Dim tooltip As New ToolTip
        tooltip.SetToolTip(btnSave, "Save")
        tooltip.SetToolTip(btnEdit, "Edit")
        tooltip.SetToolTip(btnDelete, "Delete")
        tooltip.SetToolTip(btnRefresh, "Refresh")
        tooltip.SetToolTip(btnSort, "Sort")
        tooltip.SetToolTip(btnHighlight, "Highlight")
        tooltip.SetToolTip(btnSuggest, "Suggest")
        tooltip.SetToolTip(btnFilter, "Filter")
        tooltip.SetToolTip(Button2, "View Family Schedule")

        tooltip.SetToolTip(ComboBox3, "Select the meal")
        LoadMealPlanfromDatabase1()
        PopulateDataGridView()

        ComboBox3.Items.Clear()
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            ' Query to fetch all ItemName values from Inventory1
            Dim fetchcommand As New OleDbCommand("SELECT ItemName FROM Inventory", conn)
            Using Readers As OleDbDataReader = fetchcommand.ExecuteReader()
                ' Add ItemName values to ComboBox1
                While Readers.Read()
                    ComboBox3.Items.Add(Readers("ItemName").ToString())
                End While
            End Using
        End Using
        ''hhhhhh
        PopulateComboboxFromDatabase(ComboBox4)
    End Sub

    Private mealPlanData As DataTable
    Private currentRowIndex As Integer = 0

    ' Load filtered meal plan data based on frequency
    Private Sub LoadFilteredMealPlan()
        Using dbConnection As New OleDbConnection(connectionString)
            Dim selectedFilter As String = ComboBox2.SelectedItem?.ToString()
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

    Public Sub LoadMealPlanfromDatabase1()
        Try

            Debug.WriteLine("DataGridview populated successfully MealPlanForm_Load")
            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                conn.Open()

                ' Update the table name if necessary  
                Dim tableName As String = "MealPlans"


                ' Create an OleDbCommand to select the data from the database  
                Dim cmd As New OleDbCommand($"SELECT * FROM {tableName}", conn)

                ' Create a DataAdapter and fill a DataTable  
                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable()
                da.Fill(dt)

                ' Bind the DataTable to the DataGridView  
                DataGridView1.DataSource = dt
            End Using

        Catch ex As Exception
            Debug.WriteLine($"DataGridView population failed")
            Debug.WriteLine($"Unexpected error in DataGridView: {ex.Message}")
            Debug.WriteLine($"Error in PopulateDataGridView: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            'MessageBox.Show("An error occurred while loading data into the grid.", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Property meals As New List(Of MealPlans)
    Private Sub PopulateDataGridView()

        'Add each expense to the DataGridView
        For Each meal As MealPlans In meals()

            DataGridView1.Rows.Add(meal.MealName, meal.StartDate.ToShortDateString(), meal.FilePath, meal.Description, meal.EndDate.ToShortDateString(), meal.Meals,
                                  meal.TotalCalories, meal.Calories, meal.Frequency, meal.Items)
            Try
                Debug.WriteLine("PopulateDataGridView: DataGridView populated successfully.")

            Catch ex As Exception
                Debug.WriteLine($"Error in PopulateDataGridView: {ex.Message}")
                Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
                MessageBox.Show("An error occurred while loading data into the grid.", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Next

    End Sub
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Try
            If DataGridView1.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

                ' Load the data from the selected row into UI controls  

                TextBox2.Text = selectedRow.Cells("Description").Value.ToString()
                NumericUpDown1.Text = selectedRow.Cells("TotalCalories").Value.ToString()
                DateTimePicker1.Text = selectedRow.Cells("StartDate").Value.ToString()
                DateTimePicker2.Text = selectedRow.Cells("EndDate").Value.ToString()
                TextBox3.Text = selectedRow.Cells("FilePath").Value.ToString()
                ComboBox3.SelectedItem = selectedRow.Cells("Items").Value.ToString()
                ComboBox1.SelectedItem = selectedRow.Cells("Calories").Value.ToString()
                ComboBox2.SelectedItem = selectedRow.Cells("Frequency").Value.ToString()
                lstMealSuggestions.SelectedItem = selectedRow.Cells("Meals").Value.ToString()
                TextBox4.Text = selectedRow.Cells("MealName").Value.ToString()

                ' Enable/ disable the buttons based on the selected person  
                btnSave.Enabled = False
                btnDelete.Enabled = True
                btnEdit.Enabled = True
            End If

            Debug.WriteLine("DataGridView populated successfully")


            'Calculation logic here...
        Catch ex As FormatException
            Debug.WriteLine("Select the datagridview.")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("Please Select the datagridview.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Catch ex As Exception
            Debug.WriteLine($"Unexpected error Selecting the datagridview {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("An unexpected error occured Selecting the datagridview.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Debug.WriteLine("The DataGridView selected unsuccessful.")


        If DataGridView1.SelectedRows.Count > 0 Then
            Dim filePath As String = DataGridView1.SelectedRows(0).Cells("FilePath").Value.ToString()

            If System.IO.File.Exists(filePath) Then
                PictureBox1.Image = Image.FromFile(filePath)
            Else
                PictureBox1.Image = Nothing
                MessageBox.Show("Image file not found.")
            End If
        End If


    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If DataGridView1.SelectedRows.Count > 0 Then

            ' Get the selected row  
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Assuming there is an "ID" column which is the primary key in the database  
            Dim MealPlanId As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  

            ' Confirm deletion  
            Dim confirmationResult As DialogResult = MessageBox.Show("Are you sure you want to delete this mealplan?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

            If confirmationResult = DialogResult.Yes Then
                ' Proceed with deletion  
                Try
                    Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                        conn.Open()

                        ' Create the delete command  
                        Dim cmd As New OleDbCommand("DELETE FROM [MealPlans] WHERE [ID] = ?", conn)
                        cmd.Parameters.AddWithValue("@ID", MealPlanId) ' Primary key for matching record  

                        ' Execute the delete command  
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show("meals deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ' Optionally refresh DataGridView or reload from database  

                        Else
                            MessageBox.Show("No meals was deleted. Please check if the ID exists.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End Using

                Catch ex As Exception
                    MessageBox.Show($"An error occurred while deleting the meals: {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            MessageBox.Show("Please select an meals to delete.", "deletetion error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub
    Public Function SuggestMeals() As List(Of String)
        Dim suggestedMeals As New List(Of String)

        Try
            Using connect As New OleDbConnection(connectionString)

            End Using
            conn.Open()

            ' Get all meal recipes
            Dim mealQuery As String = "SELECT ItemName,Ingredients FROM Recipes"
            Dim mealCommand As New OleDb.OleDbCommand(mealQuery, conn)
            Dim mealReader As OleDb.OleDbDataReader = mealCommand.ExecuteReader()

            While mealReader.Read()
                Dim mealName As String = mealReader("ItemName").ToString()
                Dim requiredIngredients As String() = mealReader("Ingredients").ToString().Split(",")

                Dim allIngredientsAvailable As Boolean = True

                ' Check if all required ingredients exist in GroceryInventory and are not expired
                For Each ingredient In requiredIngredients
                    Dim trimmedIngredient As String = ingredient.Trim()
                    Dim checkQuery As String = "SELECT ExpiryDate FROM GroceryItem WHERE Category AND Quantity > 0"
                    Dim checkCommand As New OleDb.OleDbCommand(checkQuery, conn)
                    checkCommand.Parameters.AddWithValue("@Ingredients", trimmedIngredient)

                    Dim expirationDate As Object = checkCommand.ExecuteScalar()

                    ' Check if the ingredient exists and its expiration date
                    If expirationDate Is Nothing Then
                        allIngredientsAvailable = False
                    Else
                        ' Validate that the ingredient is not expired
                        If Convert.ToDateTime(expirationDate) < DateTime.Now Then
                            allIngredientsAvailable = False
                        End If
                    End If

                    If Not allIngredientsAvailable Then
                        Exit For
                    End If
                Next

                ' If all ingredients are available and not expired, add the meal to suggested list
                If allIngredientsAvailable Then
                    suggestedMeals.Add(mealName)
                End If
            End While
            mealReader.Close()



        Catch ex As Exception
            MsgBox("Error suggesting meals: " & ex.Message, MsgBoxStyle.Critical, "Database Error")
        Finally
            conn.Close()
        End Try

        Return suggestedMeals
    End Function

    Private Sub btnSuggest_Click(sender As Object, e As EventArgs) Handles btnSuggest.Click

        Dim suggestedMeals As New List(Of String)
        Dim dietaryPreference As String = ""
        Dim allowedFoodTypes As New List(Of String)

        If ComboBox4.SelectedItem Is Nothing OrElse ComboBox3.SelectedItem Is Nothing Then
            MsgBox("Please select a person and an ingredient.")
            Return
        End If

        Dim selectedPerson As String = ComboBox4.SelectedItem.ToString()
        Dim selectedIngredient As String = ComboBox3.SelectedItem.ToString().ToLower()

        Try
            Using conn As New OleDbConnection(connectionString)
                conn.Open()

                ' Get person's first and last name
                Dim names = selectedPerson.Split(" "c)
                If names.Length < 2 Then
                    MsgBox("Please select a full name with First and Last name.")
                    Return
                End If

                Dim firstName As String = names(0).Trim()
                Dim lastName As String = names(1).Trim()

                ' Get dietary preference and food types
                Dim personCmd As New OleDbCommand("SELECT Dietary, FoodType FROM PersonalDetails WHERE FirstName = ? AND LastName = ?", conn)
                personCmd.Parameters.AddWithValue("?", firstName)
                personCmd.Parameters.AddWithValue("?", lastName)

                Using reader = personCmd.ExecuteReader()
                    If reader.Read() Then
                        dietaryPreference = reader("Dietary").ToString().ToLower()
                        allowedFoodTypes = reader("FoodType").ToString().ToLower().Split(","c).Select(Function(f) f.Trim()).ToList()
                        MsgBox($"{firstName} {lastName} has a '{dietaryPreference}' preference. Allowed food types: {String.Join(", ", allowedFoodTypes)}")
                    Else
                        MsgBox("No dietary info found for " & selectedPerson)
                        Return
                    End If
                End Using

                lstMealSuggestions.Items.Clear()

                ' Find meals that use the selected ingredient
                Dim mealQuery As String = "SELECT ItemName, Ingredients FROM Recipe"
                Dim ingredientUsedInMeals As New List(Of String)

                Using mealCmd As New OleDbCommand(mealQuery, conn)
                    Using mealReader As OleDbDataReader = mealCmd.ExecuteReader()
                        While mealReader.Read()
                            Dim mealName As String = mealReader("ItemName").ToString()
                            Dim ingredients As String() = mealReader("Ingredients").ToString().ToLower().Split(","c).Select(Function(x) x.Trim()).ToArray()

                            ' Only check meals that include the selected ingredient
                            If Not ingredients.Contains(selectedIngredient) Then Continue While

                            ingredientUsedInMeals.Add(mealName)

                            ' Check selected ingredient in inventory
                            Dim invCmd As New OleDbCommand("SELECT Quantity, ExpiryDate, Category FROM Inventory WHERE LCASE(TRIM(ItemName)) = ?", conn)
                            invCmd.Parameters.AddWithValue("?", selectedIngredient)

                            Using invReader = invCmd.ExecuteReader()
                                If invReader.Read() Then
                                    Dim qty As Integer = Convert.ToInt32(invReader("Quantity"))
                                    Dim expiry As Date = Convert.ToDateTime(invReader("ExpiryDate"))
                                    Dim category As String = invReader("Category").ToString().ToLower().Trim()

                                    If qty <= 0 OrElse expiry < Date.Today Then
                                        MsgBox("❌ Cannot suggest '" & mealName & "' because '" & selectedIngredient & "' is expired or out of stock.")
                                        Continue While
                                    End If

                                    If Not allowedFoodTypes.Contains(category) Then
                                        MsgBox("❌ Cannot suggest '" & mealName & "' because '" & selectedIngredient & "' (category '" & category & "') is not allowed for " & dietaryPreference & " diet.")
                                        Continue While
                                    End If
                                Else
                                    MsgBox("❌ Cannot suggest '" & mealName & "' because '" & selectedIngredient & "' is not found in inventory.")
                                    Continue While
                                End If
                            End Using

                            ' If all OK, suggest the meal
                            suggestedMeals.Add(mealName)
                        End While
                    End Using
                End Using

                ' Show suggested meals
                lstMealSuggestions.Items.Clear()
                If suggestedMeals.Count > 0 Then
                    lstMealSuggestions.Items.Add("✅ You can make these meals using '" & selectedIngredient & "':")
                    For Each meal In suggestedMeals
                        lstMealSuggestions.Items.Add(" - " & meal)
                    Next

                    ' Show message if selected ingredient is used in multiple meals
                    If ingredientUsedInMeals.Count > 1 Then
                        MsgBox("The ingredient '" & selectedIngredient & "' is used in multiple meals. You can still make: " & String.Join(", ", suggestedMeals))
                    End If
                Else
                    lstMealSuggestions.Items.Add("❌ No valid meals found with the selected ingredient.")
                End If
            End Using

        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try

    End Sub
    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        LoadMealPlanfromDatabase1()
        'Rinae.ClearControls(Me)
    End Sub

    Private Sub btnSort_Click(sender As Object, e As EventArgs) Handles btnSort.Click
        DataGridView1.Sort(DataGridView1.Columns("TotalCalories"), System.ComponentModel.ListSortDirection.Ascending)
        DataGridView1.Sort(DataGridView1.Columns("StartDate"), System.ComponentModel.ListSortDirection.Ascending)
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        Try

            Using conn As New OleDbConnection(connectionString)
                conn.Open()

                ' Update the table name if necessary  


                ' Create an OleDbCommand to insert the Expense data into the database  
                Dim cmd As New OleDbCommand("INSERT INTO MealPlans ([StartDate], [EndDate], [Meals], [MealName], [Items], [TotalCalories], [Description], [FilePath], [Calories], [Frequency]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", conn)


                Dim meals As New MealPlans

                'Assign Values 
                meals.StartDate = DateTimePicker1.Text
                meals.EndDate = DateTimePicker2.Text
                meals.Meals = lstMealSuggestions.SelectedItem.ToString()
                meals.MealName = TextBox4.Text
                meals.Items = ComboBox3.SelectedItem.ToString
                meals.TotalCalories = NumericUpDown1.Value
                meals.Description = TextBox2.Text
                meals.FilePath = TextBox3.Text
                meals.Calories = ComboBox1.SelectedItem.ToString
                meals.Frequency = ComboBox2.SelectedItem.ToString


                ' Set the parameter values from the UI controls  
                'For Each exp As Expenses In expenses
                cmd.Parameters.Clear()


                cmd.Parameters.AddWithValue("@StartDate", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@EndDate", DateTimePicker2.Text)
                cmd.Parameters.AddWithValue("@Meals", lstMealSuggestions.SelectedItem.ToString)
                cmd.Parameters.AddWithValue("@MealName", TextBox4.Text)
                cmd.Parameters.AddWithValue("@Items", ComboBox3.SelectedItem.ToString)
                cmd.Parameters.AddWithValue("@TotalCalories", NumericUpDown1.Value)
                cmd.Parameters.AddWithValue("@Description", TextBox2.Text)
                cmd.Parameters.AddWithValue("@FilePath", TextBox3.Text)
                cmd.Parameters.AddWithValue("@Calories", ComboBox1.SelectedItem.ToString)
                cmd.Parameters.AddWithValue("@Frequency", ComboBox2.SelectedItem.ToString)

                cmd.ExecuteNonQuery()

                'Display a confirmation messageBox  
                MsgBox("Schedule Information Added!" & vbCrLf &
                "StartDate: " & meals.StartDate.ToString & vbCrLf &
                "EndDate: " & meals.EndDate.ToString & vbCrLf &
                "Meals: " & meals.Meals & vbCrLf &
                "Items: " & meals.Items & vbCrLf &
                 "MealName: " & meals.MealName & vbCrLf &
                "TotalCalories: " & meals.TotalCalories & vbCrLf &
                "Description: " & meals.Description & vbCrLf &
                 "FilePath: " & meals.FilePath & vbCrLf &
                 "Calories: " & meals.Calories & vbCrLf &
                "Frequency: " & meals.Frequency.ToString(), vbInformation, "Schedule confirmation")

            End Using

        Catch ex As Exception
            Debug.WriteLine($"Database error in btnAdd_Click: {ex.Message}")
            MessageBox.Show("Error saving Schedule to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine("Exiting btnAdd")

        End Try


        '' MealPlan Form - Add this code to btnAddMeal_Click

        'Dim mealDate As Date = DateTimePicker1.Value.Date
        'Dim mealTitle As String = TextBox4.Text.Trim()

        'If mealTitle = "" Then
        '    MessageBox.Show("Please enter a meal name.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        '    Exit Sub
        'End If




        'Using conn As New OleDbConnection(connectionString)
        '    Dim query As String = "INSERT INTO FamilySchedule (DateOfEvent, Title, EventType) VALUES (?, ?, ?)"
        '    Using cmd As New OleDbCommand(query, conn)
        '        cmd.Parameters.AddWithValue("?", mealDate)
        '        cmd.Parameters.AddWithValue("?", mealTitle)
        '        cmd.Parameters.AddWithValue("?", "MealName")
        '        conn.Open()
        '        cmd.ExecuteNonQuery()
        '    End Using
        'End Using

        'MessageBox.Show("Meal added to schedule successfully!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)

        '' Optional: clear fields
        'TextBox4.Clear()
    End Sub


    Private Sub btnHighlight_Click(sender As Object, e As EventArgs) Handles btnHighlight.Click
        Try
            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.Cells("StartDate").Value IsNot Nothing Then
                    Dim StartDate As DateTime = Convert.ToDateTime(row.Cells("StartDate").Value)
                    Dim EndDate As DateTime = Convert.ToDateTime(row.Cells("EndDate").Value)
                    If StartDate <= DateTimePicker1.Value AndAlso EndDate >= DateTimePicker2.Value Then
                        row.DefaultCellStyle.BackColor = Color.Red

                    Else
                        row.DefaultCellStyle.BackColor = Color.White
                    End If
                End If
            Next
            Dim incmpleteCount As Integer = 0
            'For Each row As DataGridViewRow In DataGridView1.Rows
            '    If row.Cells("StartDate").Value IsNot Nothing AndAlso row.Cells("EndDate").Value.ToString() <> ">27, 1, 2025" Then
            incmpleteCount += 1
            '    End If
            'Next
            Label14.Text = "Incomplete MealPlan:" & incmpleteCount.ToString
        Catch ex As Exception
            MessageBox.Show("Error highlighting overdue meals")

        End Try
    End Sub

    Private Sub btnFilter_Click(sender As Object, e As EventArgs) Handles btnFilter.Click

        Dim SelectedCalories As String = If(ComboBox1.SelectedItem IsNot Nothing, ComboBox1.SelectedItem.ToString(), "")
        HouseHoldManagment_Module.FilterMealPlan(SelectedCalories)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Image Files|*.jpg;*.*.bmp"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            PictureBox1.ImageLocation = openFileDialog.FileName
            TextBox3.Text = openFileDialog.FileName
        End If
    End Sub

    Private Sub Panel2_Paint(sender As Object, e As PaintEventArgs) Handles Panel2.Paint

    End Sub
    Public Sub PopulateComboboxFromDatabase(ByRef comboBox As ComboBox)
        Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
        'Dim conn As New OleDbConnection(Rinae.connectionString)
        Try
            Debug.WriteLine("populating combobox from database successfully!")
            ' 1. Open the database connection  
            conn.Open()

            ' 2. Retrieve the FirstName and LastName columns from the PersonalDetails table  
            Dim query As String = "SELECT FirstName, LastName FROM PersonalDetails"
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            ' 3. Bind the retrieved data to the combobox  
            comboBox.Items.Clear()
            While reader.Read()
                comboBox.Items.Add($"{reader("FirstName")} {reader("LastName")}")
            End While

            ' 4. Close the database connection  
            reader.Close()

        Catch ex As Exception
            Debug.WriteLine($"form loaded unsuccessful")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
        Finally
            ' Close the database connection  
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'If MessageBox.Show("Would you like to view the calendar?", "Open Family Schedule", MessageBoxButtons.YesNo) = DialogResult.Yes Then
        '    Family_Schedule.ShowDialog()

        'End If
        If MessageBox.Show("Would you like to view the calendar?", "Open Family Schedule", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            ' Create the Family Schedule form
            Dim familyScheduleForm As New Family_Schedule()

            ' Enable auto-filtering
            familyScheduleForm.AutoFilterEnabled = True

            ' Show the form
            familyScheduleForm.ShowDialog()
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ComboBox1.SelectedItem Is Nothing Then Exit Sub

        Dim selectedIngredient As String = ComboBox1.SelectedItem.ToString().Trim().ToLower()

        Try
            Using conn As New OleDbConnection(connectionString)
                conn.Open()

                Dim cmd As New OleDbCommand("SELECT Quantity, ExpiryDate FROM Inventory WHERE LCASE(TRIM(ItemName)) = ?", conn)
                cmd.Parameters.AddWithValue("?", selectedIngredient)

                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        Dim qty As Integer = Convert.ToInt32(reader("Quantity"))
                        Dim expiryDate As Date = Convert.ToDateTime(reader("ExpiryDate"))

                        If qty <= 0 Then
                            MessageBox.Show(selectedIngredient & " is out of stock.", "Out of Stock", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            Return
                        End If

                        If expiryDate < Date.Today Then
                            MessageBox.Show(selectedIngredient & " is expired (Expiry: " & expiryDate.ToShortDateString() & ").", "Expired Item", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            Return
                        End If

                    Else
                        MessageBox.Show(selectedIngredient & " was not found in the inventory.", "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End If
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Error checking ingredient: " & ex.Message)
        End Try
    End Sub
End Class