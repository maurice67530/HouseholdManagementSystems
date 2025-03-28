
Imports System.Windows.Forms
Imports System.Data.OleDb

Public Class Budjet

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        DashBoard.Show()
        Me.Close()
    End Sub

    Private Sub Budjet_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        LoadBudgetFromDatabase()


        ' Set tooltips For buttons
        ToolTip1.SetToolTip(Button1, "Dashboard")
        ToolTip1.SetToolTip(Button2, "Save")
        ToolTip1.SetToolTip(Button3, "Edit")
        ToolTip1.SetToolTip(Button4, "Delete")
        ToolTip1.SetToolTip(Button5, "Sort")
        ToolTip1.SetToolTip(Button6, "Filter")
    End Sub
    Public Sub LoadBudgetFromDatabase()
        Try
            Debug.WriteLine("populated connected succesfully")
            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()

                Dim tableName As String = "Budget"
                Dim cmd As New OleDbCommand($"SELECT * FROM {tableName}", conn)

                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)

                DataGridView1.DataSource = dt

            End Using

        Catch ex As Exception

            Debug.WriteLine("populated fail to initialize succesfully")
            MessageBox.Show($"Error Loading  Expenses data from database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Try


            Dim Budget As New Budget With {
           .paymentMethod = ComboBox1.Text,
            .PayingList = TextBox2.Text,
           .Amount = TextBox3.Text,
           .category = TextBox4.Text,
            .Quantity = TextBox5.Text}


            Using connect As New OleDbConnection(Module1.connectionString)
                Module1.conn.Open()

                Dim tablename As String = "Budget"

                Dim cmd As New OleDbCommand($"INSERT INTO [Budget] ([paymentMethod],[payingList],[Amount],[Category],[Quantity]) VALUES (@PaymentMethod,@PayingList ,@Amount, @Category, @Quantity)", Module1.conn)
                'set the parameter values from the UI controls
                'class declaration  


                'params
                cmd.Parameters.AddWithValue("@PaymentMethod", Budget.paymentMethod)
                cmd.Parameters.AddWithValue("@PayingList", Budget.PayingList)
                cmd.Parameters.AddWithValue("@Amount", Budget.Amount)
                cmd.Parameters.AddWithValue("@category", Budget.category)
                cmd.Parameters.AddWithValue("@Quantity", Budget.Quantity)
                cmd.ExecuteNonQuery()

            End Using

            MsgBox("Task Management Added!" & vbCrLf &
                 "PaymentMethod: " & Budget.paymentMethod & vbCrLf &
                   "PayingList: " & Budget.PayingList & vbCrLf &
                   "Amount: " & Budget.Amount & vbCrLf &
                   "Category: " & Budget.category & vbCrLf &
                   "Quantity: " & Budget.Quantity, vbInformation, "Budget Confirmation")




        Catch ex As OleDbException
            Debug.WriteLine($"Database error: {ex.Message}")
            'MessageBox.Show("Error saving test to database. Please check the connection.", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Debug.WriteLine($"General error: {ex.Message}")
            'MessageBox.Show("An unexpected error occurred.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        Finally

        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Debug.WriteLine("Entering btnEdit")
            Dim PaymentMethod As String = ComboBox1.Text
            Dim Payinglist As String = TextBox2.Text
            Dim Amount As String = TextBox3.Text
            Dim category As String = TextBox3.Text
            Dim Quantity As String = TextBox4.Text



            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()

                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Create an OleDbCommand to update the expense data in the database  

                Dim cmd As New OleDbCommand("UPDATE [Budget] SET [PaymentMethod] = ?, [PayingList] = ?, [Amount] = ?, [category] = ?, [Quantity] =?WHERE [ID] = ?", conn)
                'Set the parameter values from the UI controls 


                Dim Budget As New Budget
                cmd.Parameters.AddWithValue("@PaymentMethod", Budget.paymentMethod)
                cmd.Parameters.AddWithValue("@PayingList", Budget.PayingList)
                cmd.Parameters.AddWithValue("@Amount", Budget.Amount)
                cmd.Parameters.AddWithValue("@category", Budget.category)
                cmd.Parameters.AddWithValue("@Quantity", Budget.Quantity)
                cmd.Parameters.AddWithValue("@ID", category) ' Primary key for matching record  
                cmd.ExecuteNonQuery()

                MsgBox("Budget Updated Successfuly!", vbInformation, "Update Confirmation")


            End Using
        Catch ex As OleDbException
            MessageBox.Show($"Error updating Budget in database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        conn.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' Check if there are any selected rows in the DataGridView for expenses  
        If DataGridView1.SelectedRows.Count > 0 Then
            ' Get the selected row  
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Assuming there is an "ID" column which is the primary key in the database  
            Dim Budget As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  

            ' Confirm deletion  
            Dim confirmationResult As DialogResult = MessageBox.Show("Are you sure you want to delete this expense?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If confirmationResult = DialogResult.Yes Then
                ' Proceed with deletion  
                Try
                    Using conn As New OleDbConnection(Module1.connectionString)
                        conn.Open()

                        ' Create the delete command  
                        Dim cmd As New OleDbCommand("DELETE FROM [Budget] WHERE [ID] = ?", conn)
                        cmd.Parameters.AddWithValue("@ID", Budget) ' Primary key for matching record  

                        ' Execute the delete command  
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show("MealPlan deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ' Optionally refresh DataGridView or reload from database  
                            'PopulateDataGridView()
                        Else
                            MessageBox.Show("No expense was deleted. Please check if the ID exists.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End Using

                Catch ex As Exception
                    MessageBox.Show($"An error occurred while deleting the TaskManagement: {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            MessageBox.Show("Please select an TaskManagement to delete.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            conn.Close()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
    End Sub
End Class