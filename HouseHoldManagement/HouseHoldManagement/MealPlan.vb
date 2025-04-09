Imports System.IO
Imports System.Data.OleDb

Public Class MealPlan
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
        ComboBox3.Items.AddRange(New String() {"<500", "500-1000", ">1000"})
        ComboBox1.Items.AddRange(New String() {"Day", "Week", "Month"})

        Dim tooltip As New ToolTip
        tooltip.SetToolTip(btnSave, "Save")
        tooltip.SetToolTip(btnRefresh, "Refresh")
        tooltip.SetToolTip(btnEdit, "Edit")
        tooltip.SetToolTip(btnDelete, "Delete")
        tooltip.SetToolTip(btnFilter, "Filter")
        tooltip.SetToolTip()
        tooltip.SetToolTip()
    End Sub
End Class