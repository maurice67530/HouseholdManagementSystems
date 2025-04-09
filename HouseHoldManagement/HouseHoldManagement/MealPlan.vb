Imports System.IO
Imports System.Data.OleDb

Public Class MealPlan
    Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        Try
            Debug.WriteLine("Entering btnEdit_Click")
            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                conn.Open()

                Dim tablename As String = "MealPlans"
                Using cmd As New OleDbCommand("INSERT INTO MealPlans ([StartDate], [EndDate], [Meals], [MealName], [Items], [TotalCalories], [Description], [FilePath], [Calories], [Frequency]) VALUES (@StartDate, @EndDate, @Meals, @MealName, @Items, @TotalCalories@Description, @FilePath, @Calories, @Frequency)", conn)

                    cmd.Parameters.AddWithValue("@StartDate", DateTimePicker1.Text)
                    cmd.Parameters.AddWithValue("@EndDate", DateTimePicker2.Text)
                    cmd.Parameters.AddWithValue("@Meals", ListBox1.SelectedItem.ToString)
                    cmd.Parameters.AddWithValue("@MealName", TextBox4.Text)
                    cmd.Parameters.AddWithValue("@Items", ComboBox2.SelectedItem.ToString)
                    cmd.Parameters.AddWithValue("@TotalCalories", NumericUpDown1.Text)
                    cmd.Parameters.AddWithValue("@Description", TextBox2.Text)
                    cmd.Parameters.AddWithValue("@FilePath", TextBox3.Text)
                    cmd.Parameters.AddWithValue("@Calories", ComboBox3.SelectedItem.ToString)
                    cmd.Parameters.AddWithValue("@Frequency", ComboBox1.SelectedItem.ToString)
                    cmd.ExecuteNonQuery()
                End Using
                MessageBox.Show("Edited successfully")

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

        Debug.WriteLine("Existing btnEdit_Click")
    End Sub

    Private Sub MealPlan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.AddRange(New String() {"<500", "500-1000", ">1000"})
        ComboBox2.Items.AddRange(New String() {"Day", "Week", "Month"})
        ComboBox3.Items.AddRange(New String() {"Noodles", "Chicken", "Bread"})
        ListBox1.Items.AddRange(New String() {"Noodles", "Chicken Curry", "Kota"})
        Dim tooltip As New ToolTip
        tooltip.SetToolTip(btnSave, "Save")
        tooltip.SetToolTip(btnEdit, "Edit")
        tooltip.SetToolTip(btnDelete, "Delete")
        tooltip.SetToolTip(btnRefresh, "Refresh")
        tooltip.SetToolTip(btnSort, "Sort")
        tooltip.SetToolTip(btnHighlight, "Highlight")
        tooltip.SetToolTip(btnPrint, "Print")
        tooltip.SetToolTip(btnSuggest, "Suggest")
        tooltip.SetToolTip(btnFilter, "Filter")

        PopulateDataGridView()
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
                ListBox1.SelectedItem = selectedRow.Cells("Meals").Value.ToString()
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

    '    Public Function SuggestMeals() As List(Of String)
    '        Dim suggestedMeals As New List(Of String)

    '        Try
    '            Using connect As New OleDbConnection(connectionString)

    '            End Using
    '            conn.Open()

    '            ' Get all meal recipes
    '            Dim mealQuery As String = "SELECT MealName,Items FROM MealPlans"
    '            Dim mealCommand As New OleDb.OleDbCommand(mealQuery, conn)
    '            Dim mealReader As OleDb.OleDbDataReader = mealCommand.ExecuteReader()

    '            While mealReader.Read()
    '                Dim mealName As String = mealReader("MealName").ToString()
    '                Dim requiredIngredients As String() = mealReader("Items").ToString().Split(",")

    '                Dim allIngredientsAvailable As Boolean = True

    '                ' Check if all required ingredients exist in GroceryInventory and are not expired
    '                For Each ingredient In requiredIngredients
    '                    Dim trimmedIngredient As String = ingredient.Trim()
    '                    Dim checkQuery As String = "SELECT ExpiryDate FROM GroceryItemss WHERE ItemName=@Ingredients AND Quantity > 0"
    '                    Dim checkCommand As New OleDb.OleDbCommand(checkQuery, conn)
    '                    checkCommand.Parameters.AddWithValue("@Ingredients", trimmedIngredient)

    '                    Dim expirationDate As Object = checkCommand.ExecuteScalar()

    '                    ' Check if the ingredient exists and its expiration date
    '                    If expirationDate Is Nothing Then
    '                        allIngredientsAvailable = False
    '                    Else
    '                        ' Validate that the ingredient is not expired
    '                        If Convert.ToDateTime(expirationDate) < DateTime.Now Then
    '                            allIngredientsAvailable = False
    '                        End If
    '                    End If

    '                    If Not allIngredientsAvailable Then
    '                        Exit For
    '                    End If
    '                Next

    '                ' If all ingredients are available and not expired, add the meal to suggested list
    '                If allIngredientsAvailable Then
    '                    suggestedMeals.Add(mealName)
    '                End If
    '            End While
    '            mealReader.Close()



    '        Catch ex As Exception
    '            MsgBox("Error suggesting meals: " & ex.Message, MsgBoxStyle.Critical, "Database Error")
    '        Finally
    '            conn.Close()
    '        End Try

    '        Return suggestedMeals
    '    End Function

    '    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
    '        'Module1.Mains()

    '        ListBox1.Items.Clear()
    '        TextBox2.ReadOnly = True
    '        Dim meals As List(Of String) = SuggestMeals()

    '        If meals.Count > 0 Then
    '            ForEach meal In meals
    '               ListBox2.Items.Add(meal)

    '            Next
    '            MsgBox("Meal Suggestions have been prepared with current Greocery Items.", MsgBoxStyle.Information, "No Available Meals")
    '            'FetchAlternativeMeals(SuggestMeals)
    '        Else
    '            MsgBox("No meals can be prepared with current inventory.", MsgBoxStyle.Exclamation, "No Available Meals")
    '        End If
    '        End Su
    'Private Sub btnSuggest_Click(sender As Object, e As EventArgs) Handles btnSuggest.Click

    '    End Sub
End Class