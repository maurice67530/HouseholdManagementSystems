Imports System.IO
Imports System.Data.OleDb

Public Class MealPlan
    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        Try
            Debug.WriteLine("Entering btnEdit_Click")
            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                conn.Open()

                Dim tablename As String = "MealPlans"
                Using cmd As New OleDbCommand("INSERT INTO MealPlan ([MealPlanId], [Description], [TotalCalories], [StartDate], [EndDate], [Meals], [MealName], [PicturePath], [Items], [Calories], [MealPlanPrint]) VALUES (@MealPlanId, @Description, @TotalCalories, @StartDate, @EndDate, @PicturePath, @Items, @Calories, @MealPlanPrint, @Meals, @MealName)", conn)

                    cmd.Parameters.AddWithValue("@MealPlanId", HouseHoldManagment_Module.MealPlanForm1())
                    cmd.Parameters.AddWithValue("@Description", TextBox3.Text)
                    cmd.Parameters.AddWithValue("@TotalCalories", TextBox3.Text)
                    cmd.Parameters.AddWithValue("@StartDate", DateTimePicker1.Text)
                    cmd.Parameters.AddWithValue("@EndDate", DateTimePicker2.Text)
                    ' cmd.Parameters.AddWithValue("@Picturepath", TextBox5.Text)
                    cmd.Parameters.AddWithValue("@Items", ComboBox2.SelectedItem.ToString)
                    cmd.Parameters.AddWithValue("@Calories", ComboBox1.SelectedItem.ToString)
                    'cmd.Parameters.AddWithValue("@Frequency", ComboBox4.SelectedItem.ToString)
                    ' cmd.Parameters.AddWithValue("@Meals", ListBox2.SelectedItem.ToString)
                    cmd.Parameters.AddWithValue("@MealName", TextBox4.Text)
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
End Class