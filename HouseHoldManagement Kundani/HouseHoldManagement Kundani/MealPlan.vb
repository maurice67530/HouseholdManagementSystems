
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

        PopulateComboboxFromDatabase(ComboBox4)
        ComboBox1.Items.AddRange(New String() {"<500", "500-1000", ">1000"})
        ComboBox2.Items.AddRange(New String() {"Daily", "Weekly", "Monthly"})

        'lstMealSuggestions.Items.AddRange(New String() {"Noodles", "Chicken Curry", "Kota"})
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
        'Dim suggestedMeals As New List(Of String)

        'Try
        '    Using connect As New OleDbConnection(connectionString)

        '    End Using
        '    conn.Open()

        '    ' Get all meal recipes
        '    Dim mealQuery As String = "SELECT ItemName,Ingredients FROM Recipe"
        '    Dim mealCommand As New OleDb.OleDbCommand(mealQuery, conn)
        '    Dim mealReader As OleDb.OleDbDataReader = mealCommand.ExecuteReader()

        '    While mealReader.Read()
        '        Dim mealName As String = mealReader("ItemName").ToString()
        '        Dim requiredIngredients As String() = mealReader("Ingredients").ToString().Split(",")

        '        Dim allIngredientsAvailable As Boolean = True

        '        ' Check if all required ingredients exist in GroceryInventory and are not expired
        '        For Each ingredient In requiredIngredients
        '            Dim trimmedIngredient As String = ingredient.Trim()
        '            Dim checkQuery As String = "SELECT ExpiryDate FROM Inventory WHERE Category AND Quantity > 0"
        '            Dim checkCommand As New OleDb.OleDbCommand(checkQuery, conn)
        '            checkCommand.Parameters.AddWithValue("@Ingredients", trimmedIngredient)

        '            Dim expirationDate As Object = checkCommand.ExecuteScalar()

        '            ' Check if the ingredient exists and its expiration date
        '            If expirationDate Is Nothing Then
        '                allIngredientsAvailable = False
        '            Else
        '                ' Validate that the ingredient is not expired
        '                If Convert.ToDateTime(expirationDate) < DateTime.Now Then
        '                    allIngredientsAvailable = False
        '                End If
        '            End If

        '            If Not allIngredientsAvailable Then
        '                Exit For
        '            End If
        '        Next

        '        ' If all ingredients are available and not expired, add the meal to suggested list
        '        If allIngredientsAvailable Then
        '            suggestedMeals.Add(mealName)
        '        End If
        '    End While
        '    mealReader.Close()



        'Catch ex As Exception
        '    MsgBox("Error suggesting meals: " & ex.Message, MsgBoxStyle.Critical, "Database Error")
        'Finally
        '    conn.Close()
        'End Try

        'Return suggestedMeals


        ''''''



        'Dim suggestedMeals As New List(Of String)
        'Dim missingIngredients As New List(Of String)

        'Try
        '    Using conn As New OleDbConnection(connectionString)
        '        conn.Open()

        '        Dim mealQuery As String = "SELECT ItemName, Ingredients FROM Recipe"
        '        Using mealCommand As New OleDbCommand(mealQuery, conn)
        '            Using mealReader As OleDbDataReader = mealCommand.ExecuteReader()
        '                While mealReader.Read()
        '                    Dim mealName As String = mealReader("ItemName").ToString()
        '                    Dim requiredIngredients As String() = mealReader("Ingredients").ToString().Split(","c)
        '                    Dim allIngredientsAvailable As Boolean = True
        '                    Dim tempMissing As New List(Of String)

        '                    ' Check all ingredients
        '                    For Each ingredient In requiredIngredients
        '                        Dim trimmedIngredient As String = ingredient.Trim()

        '                        Dim checkQuery As String = "SELECT Quantity, ExpiryDate FROM Inventory WHERE ItemName = ?"
        '                        Using checkCommand As New OleDbCommand(checkQuery, conn)
        '                            checkCommand.Parameters.AddWithValue("?", trimmedIngredient)
        '                            Using reader As OleDbDataReader = checkCommand.ExecuteReader()
        '                                If reader.Read() Then
        '                                    Dim quantity As Integer = Convert.ToInt32(reader("Quantity"))
        '                                    Dim expiry As Date = Convert.ToDateTime(reader("ExpiryDate"))

        '                                    If quantity <= 0 Or expiry < Date.Today Then
        '                                        allIngredientsAvailable = False
        '                                        tempMissing.Add(trimmedIngredient & " (expired or out of stock)")
        '                                        Exit For
        '                                    End If
        '                                Else
        '                                    allIngredientsAvailable = False
        '                                    tempMissing.Add(trimmedIngredient & " (not found)")
        '                                    Exit For
        '                                End If
        '                            End Using
        '                        End Using
        '                    Next

        '                    ' If available, subtract quantity and add to list
        '                    If allIngredientsAvailable Then
        '                        suggestedMeals.Add(mealName)

        '                        ' Subtract 1 from each ingredient
        '                        For Each ingredient In requiredIngredients
        '                            Dim trimmedIngredient As String = ingredient.Trim()

        '                            Dim updateQuery As String = "UPDATE Inventory SET Quantity = Quantity - 1 WHERE ItemName = ?"
        '                            Using updateCommand As New OleDbCommand(updateQuery, conn)
        '                                updateCommand.Parameters.AddWithValue("?", trimmedIngredient)
        '                                updateCommand.ExecuteNonQuery()
        '                            End Using
        '                        Next
        '                    Else
        '                        missingIngredients.AddRange(tempMissing)
        '                    End If
        '                End While
        '            End Using
        '        End Using
        '    End Using

        '    ' Show missing ingredients if no meals suggested
        '    If suggestedMeals.Count = 0 AndAlso missingIngredients.Count > 0 Then
        '        lstMealSuggestions.Items.Add("No meals can be prepared.")
        '        lstMealSuggestions.Items.Add("Missing or expired ingredients:")
        '        For Each item In missingIngredients.Distinct()
        '            lstMealSuggestions.Items.Add(" - " & item)
        '        Next
        '    End If

        'Catch ex As Exception
        '    MsgBox("Error suggesting meals: " & ex.Message, MsgBoxStyle.Critical, "Database Error")
        'End Try

        'Return suggestedMeals




        ''''''



        ''Public Function SuggestMeals() As List(Of String)
        'Dim suggestedMeals As New List(Of String)
        'Dim missingIngredients As New List(Of String)

        'Try
        '    Using conn As New OleDbConnection(connectionString)
        '        conn.Open()

        '        Dim mealQuery As String = "SELECT ItemName, Ingredients FROM Recipe"
        '        Using mealCommand As New OleDbCommand(mealQuery, conn)
        '            Using mealReader As OleDbDataReader = mealCommand.ExecuteReader()
        '                While mealReader.Read()
        '                    Dim mealName As String = mealReader("ItemName").ToString()
        '                    Dim ingredients As String() = mealReader("Ingredients").ToString().Split(","c)
        '                    Dim allAvailable As Boolean = True
        '                    Dim tempMissing As New List(Of String)

        '                    ' Check each ingredient
        '                    For Each ing In ingredients
        '                        Dim ingredient = ing.Trim()
        '                        Dim checkCmd As New OleDbCommand("SELECT Quantity, ExpiryDate FROM Inventory WHERE ItemName = ?", conn)
        '                        checkCmd.Parameters.AddWithValue("?", ingredient)

        '                        Using reader = checkCmd.ExecuteReader()
        '                            If reader.Read() Then
        '                                Dim qty As Integer = Convert.ToInt32(reader("Quantity"))
        '                                Dim expiry As Date = Convert.ToDateTime(reader("ExpiryDate"))

        '                                If qty <= 0 Or expiry < Date.Today Then
        '                                    allAvailable = False
        '                                    tempMissing.Add(ingredient & " (expired or no stock)")
        '                                    Exit For
        '                                End If
        '                            Else
        '                                allAvailable = False
        '                                tempMissing.Add(ingredient & " (not found)")
        '                                Exit For
        '                            End If
        '                        End Using
        '                    Next

        '                    ' If all ingredients are fine, add meal and subtract quantities
        '                    If allAvailable Then
        '                        suggestedMeals.Add(mealName)

        '                        For Each ing In ingredients
        '                            Dim ingredient = ing.Trim()
        '                            Dim updateCmd As New OleDbCommand("UPDATE Inventory SET Quantity = Quantity - 1 WHERE ItemName = ?", conn)
        '                            updateCmd.Parameters.AddWithValue("?", ingredient)
        '                            updateCmd.ExecuteNonQuery()
        '                        Next
        '                    Else
        '                        missingIngredients.AddRange(tempMissing)
        '                    End If
        '                End While
        '            End Using
        '        End Using
        '    End Using

        '    ' Show message if none found
        '    If suggestedMeals.Count = 0 AndAlso missingIngredients.Count > 0 Then
        '        lstMealSuggestions.Items.Add("No meals can be prepared.")
        '        For Each m In missingIngredients.Distinct()
        '            lstMealSuggestions.Items.Add(" - " & m)
        '        Next
        '    End If

        'Catch ex As Exception
        '    MsgBox("Error: " & ex.Message)
        'End Try

        'Return suggestedMeals








        'hhhhh



        'Dim suggestedMeals As New List(Of String)
        'Dim missingIngredients As New List(Of String)
        'Dim dietaryPreference As String = ""
        'Dim selectedPerson As String = ComboBox4.SelectedItem.ToString()

        'Try
        '    Using conn As New OleDbConnection(connectionString)
        '        conn.Open()

        '        ' Get dietary preference
        '        Dim dietCmd As New OleDbCommand("SELECT Dietary FROM PersonalDetails WHERE FirstName = ?", conn)
        '        dietCmd.Parameters.AddWithValue("?", selectedPerson)
        '        Dim result = dietCmd.ExecuteScalar()
        '        If result IsNot Nothing Then
        '            dietaryPreference = result.ToString().ToLower()
        '        End If

        '        ' Meal suggestions
        '        Dim mealQuery As String = "SELECT ItemName, Ingredients FROM Recipe"
        '        Using mealCommand As New OleDbCommand(mealQuery, conn)
        '            Using mealReader As OleDbDataReader = mealCommand.ExecuteReader()
        '                While mealReader.Read()
        '                    Dim mealName As String = mealReader("ItemName").ToString()
        '                    Dim ingredients As String() = mealReader("Ingredients").ToString().Split(","c)
        '                    Dim allAvailable As Boolean = True
        '                    Dim tempMissing As New List(Of String)

        '                    ' Dietary check
        '                    Dim mealContainsMeat As Boolean = ingredients.Any(Function(ing) ing.ToLower().Contains("chicken") OrElse ing.ToLower().Contains("beef") OrElse ing.ToLower().Contains("fish") OrElse ing.ToLower().Contains("pork") OrElse ing.ToLower().Contains("wors") OrElse ing.ToLower().Contains("sausage"))

        '                    If dietaryPreference = "vegan" AndAlso mealContainsMeat Then
        '                        MsgBox(selectedPerson & " is vegan and does not eat " & mealName)
        '                        Return suggestedMeals
        '                    End If

        '                    If dietaryPreference = "vegetarian" AndAlso mealContainsMeat Then
        '                        MsgBox(selectedPerson & " is vegetarian and does not eat " & mealName)
        '                        Return suggestedMeals
        '                    End If

        '                    ' Check each ingredient
        '                    For Each ing In ingredients
        '                        Dim ingredient = ing.Trim()
        '                        Dim checkCmd As New OleDbCommand("SELECT Quantity, ExpiryDate FROM Inventory WHERE ItemName = ?", conn)
        '                        checkCmd.Parameters.AddWithValue("?", ingredient)
        '                        Using reader = checkCmd.ExecuteReader()
        '                            If reader.Read() Then
        '                                Dim qty As Integer = Convert.ToInt32(reader("Quantity"))
        '                                Dim expiry As Date = Convert.ToDateTime(reader("ExpiryDate"))
        '                                If qty <= 0 Or expiry < Date.Today Then
        '                                    allAvailable = False
        '                                    tempMissing.Add(ingredient & " (expired or no stock)")
        '                                    Exit For
        '                                End If
        '                            Else
        '                                allAvailable = False
        '                                tempMissing.Add(ingredient & " (not found)")
        '                                Exit For
        '                            End If
        '                        End Using
        '                    Next

        '                    ' If all ingredients are fine, add meal and subtract quantities
        '                    If allAvailable Then
        '                        suggestedMeals.Add(mealName)
        '                        For Each ing In ingredients
        '                            Dim ingredient = ing.Trim()
        '                            Dim updateCmd As New OleDbCommand("UPDATE Inventory SET Quantity = Quantity - 1 WHERE ItemName = ?", conn)
        '                            updateCmd.Parameters.AddWithValue("?", ingredient)
        '                            updateCmd.ExecuteNonQuery()
        '                        Next
        '                    Else
        '                        missingIngredients.AddRange(tempMissing)
        '                    End If
        '                End While
        '            End Using
        '        End Using
        '    End Using

        '    ' Show message if none found
        '    If suggestedMeals.Count = 0 AndAlso missingIngredients.Count > 0 Then
        '        lstMealSuggestions.Items.Add("No meals can be prepared.")
        '        For Each m In missingIngredients.Distinct()
        '            lstMealSuggestions.Items.Add(" - " & m)
        '        Next
        '    End If

        'Catch ex As Exception
        '    MsgBox("Error: " & ex.Message)
        'End Try

        'Return suggestedMeals


        'Dim suggestedMeals As New List(Of String)
        'Dim selectedItem As String = ComboBox3.SelectedItem.ToString().ToLower()

        'Try
        '    Using conn As New OleDbConnection(connectionString)
        '        conn.Open()

        '        Dim mealQuery As String = "SELECT ItemName, Ingredients FROM Recipe"
        '        Using mealCommand As New OleDbCommand(mealQuery, conn)
        '            Using mealReader As OleDbDataReader = mealCommand.ExecuteReader()
        '                While mealReader.Read()
        '                    Dim mealName As String = mealReader("ItemName").ToString()
        '                    Dim ingredients As String() = mealReader("Ingredients").ToString().Split(","c).Select(Function(i) i.Trim().ToLower()).ToArray()

        '                    If Not ingredients.Contains(selectedItem) OrElse ingredients.Length <= 1 Then
        '                        Continue While
        '                    End If

        '                    Dim allAvailable As Boolean = True

        '                    For Each ingredient In ingredients
        '                        Dim checkCmd As New OleDbCommand("SELECT Quantity, ExpiryDate FROM Inventory WHERE ItemName = ?", conn)
        '                        checkCmd.Parameters.AddWithValue("?", ingredient)

        '                        Using reader = checkCmd.ExecuteReader()
        '                            If reader.Read() Then
        '                                Dim qty As Integer = Convert.ToInt32(reader("Quantity"))
        '                                Dim expiry As Date = Convert.ToDateTime(reader("ExpiryDate"))

        '                                If expiry < Date.Today Then
        '                                    MsgBox("Item '" & ingredient & "' is expired.", MsgBoxStyle.Exclamation)
        '                                    allAvailable = False
        '                                    Exit For
        '                                ElseIf qty <= 0 Then
        '                                    MsgBox("Item '" & ingredient & "' has 0 stock.", MsgBoxStyle.Critical)
        '                                    allAvailable = False
        '                                    Exit For
        '                                ElseIf qty < 3 Then
        '                                    MsgBox("Alert: Item '" & ingredient & "' is running low (less than 3 in stock).", MsgBoxStyle.Information)
        '                                End If
        '                            Else
        '                                allAvailable = False
        '                                Exit For
        '                            End If
        '                        End Using
        '                    Next

        '                    If allAvailable Then
        '                        suggestedMeals.Add(mealName)
        '                    End If
        '                End While
        '            End Using
        '        End Using
        '    End Using

        '    lstMealSuggestions.Items.Clear()
        '    If suggestedMeals.Count = 0 Then
        '        lstMealSuggestions.Items.Add("No meals can be prepared with '" & ComboBox3.SelectedItem.ToString() & "'.")
        '    Else
        '        For Each meal In suggestedMeals
        '            lstMealSuggestions.Items.Add("✓ " & meal)
        '        Next
        '    End If

        'Catch ex As Exception
        '    MsgBox("Error: " & ex.Message)
        'End Try

        'Return suggestedMeals





        'Dim suggestedMeals As New List(Of String)
        'Dim selectedItem As String = ComboBox4.SelectedItem.ToString().ToLower()
        'Dim selectedPerson As String = ComboBox3.SelectedItem.ToString()

        'Try
        '    Using conn As New OleDbConnection(connectionString)
        '        conn.Open()

        '        ' Step 1: Check if the person is vegan
        '        Dim isVegan As Boolean = False
        '        Dim prefCmd As New OleDbCommand("SELECT Diatary FROM PersonalDetails WHERE FirstName = ?", conn)
        '        prefCmd.Parameters.AddWithValue("?", selectedPerson)
        '        Dim prefResult As Object = prefCmd.ExecuteScalar()
        '        If prefResult IsNot Nothing AndAlso prefResult.ToString().ToLower() = "vegan" Then
        '            isVegan = True
        '        End If

        '        ' Step 2: Load meat items from Inventory
        '        Dim meatItems As New List(Of String)
        '        Dim meatCmd As New OleDbCommand("SELECT ItemName FROM Inventory WHERE Category = 'Meat'", conn)
        '        Using meatReader As OleDbDataReader = meatCmd.ExecuteReader()
        '            While meatReader.Read()
        '                meatItems.Add(meatReader("ItemName").ToString().ToLower())
        '            End While
        '        End Using

        '        ' Step 3: Load meals from Recipe table
        '        Dim mealQuery As String = "SELECT ItemName, Ingredients FROM Recipe"
        '        Using mealCommand As New OleDbCommand(mealQuery, conn)
        '            Using mealReader As OleDbDataReader = mealCommand.ExecuteReader()
        '                While mealReader.Read()
        '                    Dim mealName As String = mealReader("ItemName").ToString()
        '                    Dim ingredients As String() = mealReader("Ingredients").ToString().Split(","c).Select(Function(i) i.Trim().ToLower()).ToArray()

        '                    If Not ingredients.Contains(selectedItem) OrElse ingredients.Length <= 1 Then Continue While

        '                    ' Step 4: Skip meal if vegan and it contains meat
        '                    If isVegan AndAlso ingredients.Any(Function(i) meatItems.Contains(i)) Then
        '                        Continue While
        '                    End If

        '                    ' Step 5: Check all ingredients availability
        '                    Dim allAvailable As Boolean = True
        '                    For Each ingredient In ingredients
        '                        Dim checkCmd As New OleDbCommand("SELECT Quantity, ExpiryDate FROM Inventory WHERE ItemName = ?", conn)
        '                        checkCmd.Parameters.AddWithValue("?", ingredient)
        '                        Using reader = checkCmd.ExecuteReader()
        '                            If reader.Read() Then
        '                                Dim qty As Integer = Convert.ToInt32(reader("Quantity"))
        '                                Dim expiry As Date = Convert.ToDateTime(reader("ExpiryDate"))

        '                                If expiry < Date.Today Then
        '                                    MsgBox("Item '" & ingredient & "' is expired.", MsgBoxStyle.Exclamation)
        '                                    allAvailable = False
        '                                    Exit For
        '                                ElseIf qty <= 0 Then
        '                                    MsgBox("Item '" & ingredient & "' has 0 stock.", MsgBoxStyle.Critical)
        '                                    allAvailable = False
        '                                    Exit For
        '                                ElseIf qty < 3 Then
        '                                    MsgBox("Alert: Item '" & ingredient & "' is running low.", MsgBoxStyle.Information)
        '                                End If
        '                            Else
        '                                allAvailable = False
        '                                Exit For
        '                            End If
        '                        End Using
        '                    Next

        '                    If allAvailable Then
        '                        suggestedMeals.Add(mealName)
        '                    End If
        '                End While
        '            End Using
        '        End Using
        '    End Using

        '    ' Show suggestions
        '    lstMealSuggestions.Items.Clear()
        '    If suggestedMeals.Count = 0 Then
        '        lstMealSuggestions.Items.Add("No meals can be prepared with '" & ComboBox3.SelectedItem.ToString() & "'.")
        '    Else
        '        For Each meal In suggestedMeals
        '            lstMealSuggestions.Items.Add("✓ " & meal)
        '        Next
        '    End If

        'Catch ex As Exception
        '    MsgBox("Error: " & ex.Message)
        'End Try

        'Return suggestedMeals


        '''''''''''''''''''''''''''''''


        Dim suggestedMeals As New List(Of String)
        Dim selectedItem As String = ComboBox3.SelectedItem.ToString().ToLower()
        Dim selectedPerson As String = ComboBox4.SelectedItem.ToString()

        Try
            Using conn As New OleDbConnection(connectionString)
                conn.Open()

                ' Step 1: Check if the person is vegan
                Dim isVegan As Boolean = False
                Dim prefCmd As New OleDbCommand("SELECT Dietary FROM PersonalDetails WHERE FirstName = ?", conn)
                prefCmd.Parameters.AddWithValue("?", selectedPerson)
                Dim prefResult As Object = prefCmd.ExecuteScalar()
                If prefResult IsNot Nothing AndAlso prefResult.ToString().ToLower() = "vegan" Then
                    isVegan = True
                End If

                ' Step 2: Load meat items from Inventory
                Dim meatItems As New List(Of String)
                Dim meatCmd As New OleDbCommand("SELECT ItemName FROM Inventory WHERE Category = 'Meat'", conn)
                Using meatReader As OleDbDataReader = meatCmd.ExecuteReader()
                    While meatReader.Read()
                        meatItems.Add(meatReader("ItemName").ToString().ToLower())
                    End While
                End Using

                ' Step 3: Load meals from Recipe table
                Dim mealQuery As String = "SELECT ItemName, Ingredients FROM Recipe"
                Using mealCommand As New OleDbCommand(mealQuery, conn)
                    Using mealReader As OleDbDataReader = mealCommand.ExecuteReader()
                        While mealReader.Read()
                            Dim mealName As String = mealReader("ItemName").ToString()
                            Dim ingredients As String() = mealReader("Ingredients").ToString().Split(","c).Select(Function(i) i.Trim().ToLower()).ToArray()

                            If Not ingredients.Contains(selectedItem) OrElse ingredients.Length <= 1 Then Continue While

                            ' Step 4: Skip meal if vegan and it contains meat
                            If isVegan AndAlso ingredients.Any(Function(i) meatItems.Contains(i)) Then
                                Continue While
                            End If

                            ' Step 5: Check all ingredients availability
                            Dim allAvailable As Boolean = True
                            For Each ingredient In ingredients
                                Dim checkCmd As New OleDbCommand("SELECT Quantity, ExpiryDate FROM Inventory WHERE ItemName = ?", conn)
                                checkCmd.Parameters.AddWithValue("?", ingredient)
                                Using reader = checkCmd.ExecuteReader()
                                    If reader.Read() Then
                                        Dim qty As Integer = Convert.ToInt32(reader("Quantity"))
                                        Dim expiry As Date = Convert.ToDateTime(reader("ExpiryDate"))

                                        If expiry < Date.Today Then
                                            MsgBox("Item '" & ingredient & "' is expired.", MsgBoxStyle.Exclamation)
                                            allAvailable = False
                                            Exit For
                                        ElseIf qty <= 0 Then
                                            MsgBox("Item '" & ingredient & "' has 0 stock.", MsgBoxStyle.Critical)
                                            allAvailable = False
                                            Exit For
                                        ElseIf qty < 3 Then
                                            MsgBox("Alert: Item '" & ingredient & "' is running low.", MsgBoxStyle.Information)
                                        End If
                                    Else
                                        allAvailable = False
                                        Exit For
                                    End If
                                End Using
                            Next

                            If allAvailable Then
                                suggestedMeals.Add(mealName)
                            End If
                        End While
                    End Using
                End Using
            End Using

            ' Show suggestions
            lstMealSuggestions.Items.Clear()
            If suggestedMeals.Count = 0 Then
                lstMealSuggestions.Items.Add("No meals can be prepared with '" & ComboBox3.SelectedItem.ToString() & "'.")
            Else
                For Each meal In suggestedMeals
                    lstMealSuggestions.Items.Add("✓ " & meal)
                Next
            End If

        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try

        Return suggestedMeals







    End Function


    Private Sub btnSuggest_Click(sender As Object, e As EventArgs) Handles btnSuggest.Click

        ''Module1.Mains()

        'lstMealSuggestions.Items.Clear()
        'TextBox4.ReadOnly = True
        'Dim meals As List(Of String) = SuggestMeals()

        'If meals.Count > 0 Then
        '    For Each meal In meals
        '        lstMealSuggestions.Items.Add(meal)

        '    Next
        '    MsgBox("Meal Suggestions have been prepared with current Grocery Items.", MsgBoxStyle.Information, "No Available Meals")
        '    'FetchAlternativeMeals(SuggestMeals)
        'Else
        '    MsgBox("No meals can be prepared with current inventory.", MsgBoxStyle.Exclamation, "No Available Meals")
        'End If



        lstMealSuggestions.Items.Clear()

        Dim meals As List(Of String) = SuggestMeals()

        If meals.Count > 0 Then
            lstMealSuggestions.Items.Add("Meals you can prepare:")
            For Each meal In meals
                lstMealSuggestions.Items.Add(" - " & meal)
            Next
        End If





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
    'Public Sub PopulateComboboxFromDatabase(ByRef comboBox As ComboBox)
    '    Dim conn As New OleDbConnection(connectionString)
    '    Try
    '        Debug.WriteLine("Populating combobox: combobox populated from database")
    '        'open the database connection
    '        conn.Open()

    '        ''retrieve the firstname and surname columns from the personaldetails tabel
    '        'Dim query As String = "SELECT Preference FROM Users"
    '        'Dim cmd As New OleDbCommand(query, conn)
    '        'Dim reader As OleDbDataReader = cmd.ExecuteReader()

    '        ''bind the retrieved data to the combobox
    '        'ComboBox4.Items.Clear()
    '        'While reader.Read()
    '        '    ComboBox4.Items.Add($"{reader("Preference")}")
    '        'End While

    '        ''close the database
    '        'reader.Close()

    '    Catch ex As Exception
    '        Debug.WriteLine("Failed to initialize combobox")
    '        Debug.Write($"Stack Trace: {ex.StackTrace}")
    '        'handle any exeptions that may occur
    '        MessageBox.Show($"Error: {ex.Message}")
    '    Finally
    '        'close the database connection
    '        If conn.State = ConnectionState.Open Then
    '            conn.Close()
    '        End If
    '    End Try
    '    Debug.WriteLine("Done with populating combobox from database")
    'End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'If MessageBox.Show("Would you like to view the calendar?", "Open Family Schedule", MessageBoxButtons.YesNo) = DialogResult.Yes Then
        '    Family_Schedule.ShowDialog()
        '    Family_Schedule.LoadFamilySchedule()
        '    'Family_Schedule.HighlightMealEvents()
        '    Family_Schedule.LoadFamilySchedules()

        '    Family_Schedule.btnFilte.Visible = True
        'End If

    End Sub
    Public Sub PopulateComboboxFromDatabase(ByRef comboBox As ComboBox)
        '    Dim conn As New OleDbConnection(connectionString)
        '    Try
        '        Debug.WriteLine("populate combobox successful")
        '        'open the database connection
        '        conn.Open()

        '        'retrieve the firstname and surname columns from the personaldetails tabel
        '        Dim query As String = "SELECT FirstName, LastName FROM PersonalDetails"
        '        Dim cmd As New OleDbCommand(query, conn)
        '        Dim reader As OleDbDataReader = cmd.ExecuteReader()

        '        'bind the retrieved data to the combobox
        '        ComboBox4.Items.Clear()
        '        While reader.Read()
        '            ComboBox4.Items.Add($"{reader("FirstName")} {reader("LastName")}")
        '        End While

        '        'close the database
        '        reader.Close()

        '    Catch ex As Exception
        '        'handle any exeptions that may occur  
        '        Debug.WriteLine("failed to populate combobox")
        '        Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
        '        MessageBox.Show($"Error: {ex.StackTrace}")
        '    Finally
        '        'close the database connection
        '        If conn.State = ConnectionState.Open Then
        '            conn.Close()
        '        End If
        '    End Try
        'End Sub




        Try
            Debug.WriteLine("populate combobox successful")
            conn.Open()

            ' Include DietType in the query
            Dim query As String = "SELECT FirstName, LastName, Dietary FROM PersonalDetails"
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            ComboBox4.Items.Clear()

            While reader.Read()
                Dim fullName As String = $"{reader("FirstName")} {reader("LastName")}"
                Dim dietType As String = reader("Dietary").ToString()
                ComboBox4.Items.Add($"{fullName} - {dietType}")
            End While

            reader.Close()

        Catch ex As Exception
            Debug.WriteLine("failed to populate combobox")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error: {ex.StackTrace}")

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        If ComboBox4.SelectedItem IsNot Nothing Then
            ' Split the selected item by the hyphen separator
            Dim selectedText As String = ComboBox4.SelectedItem.ToString()
            Dim parts() As String = selectedText.Split("-"c)

            If parts.Length = 2 Then
                Dim fullName As String = parts(0).Trim()
                Dim dietary As String = parts(1).Trim()

                'MessageBox.Show($"{fullName}'s is a {dietary}", "Diet Info")
            End If
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

    End Sub
End Class