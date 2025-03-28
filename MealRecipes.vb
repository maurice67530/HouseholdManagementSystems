Public Class MealRecipes
    Private ListBox1 As Object
    Private lstSuggestedMeals As Object

    'Private Sub Button1_Click(sender As Object, e As EventArgs)


    '    lstSuggestedMeals.Items.Clear()

    '    Dim meals As List(Of String) = SuggestMeals()

    '    If meals.Count > 0 Then
    '        For Each meal In meals
    '            lstSuggestedMeals.Items.Add(meal)
    '        Next
    '    Else
    '        MsgBox("No meals can be prepared with current inventory.", MsgBoxStyle.Exclamation, "No Available Meals")
    '    End If


    'End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub MealRecipes_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If ListBox1.SelectedItem Is Nothing Then
            MsgBox("Please select a meal to save.", MsgBoxStyle.Exclamation, "Select Meal")
            Exit Sub
        End If

        Dim selectedMeal As String = ListBox1.SelectedItem.ToString()

        Try
            conn.Open()

            Dim query As String = "INSERT INTO MealPlan (NameOfMeal, StartDate, EndDate) VALUES (@MealName, @StartDate, @EndDate)"
            Dim command As New OleDb.OleDbCommand(query, conn)

            command.Parameters.AddWithValue("@MealName", selectedMeal)
            command.Parameters.AddWithValue("@StartDate", Date.Today)
            command.Parameters.AddWithValue("@EndDate", Date.Today.AddDays(1)) ' One-day plan

            command.ExecuteNonQuery()
            MsgBox("Meal Plan Saved!", MsgBoxStyle.Information, "Success")

        Catch ex As Exception
            MsgBox("Error saving meal plan: " & ex.Message, MsgBoxStyle.Critical, "Database Error")
        Finally
            conn.Close()
        End Try
    End Sub
End Class