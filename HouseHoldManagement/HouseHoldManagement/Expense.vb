Imports System.IO
Imports System.Net.Mail
Imports System.Net
Imports System.Data.OleDb
Public Class Expense
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\khodani\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"
    Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)

    ' Create a ToolTip object
    Private toolTip As New ToolTip()
    Private toolTip1 As New ToolTip()
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Debug.WriteLine("Entering btnSubmit")


        Try
            'Dim txtRecentUpdate As New TextBox()
            'txtRecentUpdate.Text = $" Expense Saved at {DateTime.Now:HH:MM:ss}"
            'txtRecentUpdate.Location = New Point(65, Label20.Top + 25)
            'txtRecentUpdate.AutoSize = False
            'txtRecentUpdate.Font = New Font("Microsoft Sans Serif", 9, FontStyle.Regular)
            'txtRecentUpdate.ForeColor = Color.Black
            'Me.Controls.Add(txtRecentUpdate)

            Debug.WriteLine("User confirmed btnSubmit")
            Using conn As New OleDbConnection(connectionString)
                conn.Open()

                ' Update the table name if necessary  
                Dim tableName As String = "Expense"

                ' Create an OleDbCommand to insert the Expense data into the database 
                Dim cmd As New OleDbCommand("INSERT INTO [Expense] ([Amount], [TotalIncome], [Description], [Tags], [Currency], [Category], [Paymentmethod], [Frequency], [ApprovalStatus], [DateOfexpenses]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", conn)

                ' Set the parameter values from the UI controls 
                'Class declaretions

                Dim expense As New Expensetracking With {
                    .Amount = TextBox2.Text,
                    .TotalIncome = TextBox3.Text,
                    .Description = TextBox6.Text,
                    .Tags = TextBox4.Text,
                    .Currency = ComboBox2.SelectedItem.ToString,
                    .Category = TextBox5.Text,
                    .Paymentmethod = ComboBox1.SelectedItem.ToString,
                .Frequency = ComboBox5.SelectedItem.ToString(),
                    .ApprovalStatus = ComboBox4.SelectedItem.ToString(),
                    .DateOfexpenses = DateTimePicker1.Value}

                'txtRecentUpdate.Text = $" Expense updated at {DateTime.Now:HH:MM}"

                cmd.Parameters.Clear()


                'cmd.Parameters.AddWithValue("@ExpenseID", expense.ExpenseID)
                cmd.Parameters.AddWithValue("@Amount", expense.Amount)
                cmd.Parameters.AddWithValue("@TotalIncome", expense.TotalIncome)
                cmd.Parameters.AddWithValue("@Description", expense.Description)
                cmd.Parameters.AddWithValue("@Tags", expense.Tags)
                cmd.Parameters.AddWithValue("@Currency", expense.Currency)
                cmd.Parameters.AddWithValue("@Category", expense.Category)
                cmd.Parameters.AddWithValue("@PaymentMethod", expense.PaymentMethod)
                cmd.Parameters.AddWithValue("@Frequency", expense.Frequency)
                cmd.Parameters.AddWithValue("@ApprovalStatus", expense.ApprovalStatus)
                'cmd.Parameters.AddWithValue("@Receiver", expense.Receiver)
                cmd.Parameters.AddWithValue("@DateOfexpenses", expense.DateOfexpenses)


                MsgBox("Expense Information Saved!" & vbCrLf &
                        "ExpenseID: " & expense.ExpenseID & vbCrLf &
                        "Amount: " & expense.Amount & vbCrLf &
                        "TotalBudget: " & expense.TotalIncome & vbCrLf &
                        "Description: " & expense.Description & vbCrLf &
                        "Tags: " & expense.Tags & vbCrLf &
                        "Currency: " & expense.Currency & vbCrLf &
                        "Category: " & expense.Category & vbCrLf &
                        "PaymentMethod: " & expense.Paymentmethod & vbCrLf &
                         "Frequency: " & expense.Frequency & vbCrLf &
                         "ApprovalStatus: " & expense.ApprovalStatus & vbCrLf &
                          "Receiver: " & expense.Receiver & vbCrLf &
                          "DateOfExpense: " & expense.DateOfexpenses.ToString, vbInformation, "Expense Confirmation")

                ' Execute the SQL command to insert the data 
                ' Log the SQL statement and parameter values  

                'MessageBox.Show("Expense information saved to Database successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' Execute the SQL command to insert the data  
                cmd.ExecuteNonQuery()
                'Next

            End Using
        Catch ex As OleDbException
            'MessageBox.Show($"Error saving Expense to database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Debug.WriteLine($" Database error in Button_Click: {ex.Message}")
            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            MessageBox.Show("Error saving Expense to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            'MessageBox.Show("Error saving Expense to database.Please Check the connection.", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            'MessageBox.Show("Unexpected Error: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine($" General error in Button: {ex.Message}")
            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            Debug.WriteLine(" Failed to save")
            MessageBox.Show("An Unexpected error occured.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        conn.Close()
        Debug.WriteLine("Exiting btnSubmit")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Debug.WriteLine("Entering btnEdit")

        ' Ensure a row is selected in the DataGridView  
        If DataGridView1.SelectedRows.Count = 0 Then

            Debug.WriteLine("User canceled btnEdit")
            MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try

            'txtRecentUpdate.Text = $" Expense updated at {DateTime.Now:HH:MM}"
            'txtRecentUpdate.Location = New Point(65, Label20.Top + 45)
            'txtRecentUpdate.AutoSize = False
            'txtRecentUpdate.Font = New Font("Microsoft Sans Serif", 9, FontStyle.Regular)
            'txtRecentUpdate.ForeColor = Color.Black
            'Me.Controls.Add(txtRecentUpdate)

            'TextBox8.Text = $" Expense updated at {DateTime.Now:HH:MM}"
            'TextBox9.Text = $" Expense updated at {DateTime.Now:HH:MM}"


            Debug.WriteLine("User confirmed btnEdit")
            'Dim expenseId As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  
            Using connect As New OleDbConnection(connectionString)
                connect.Open()

                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  

                'Dim ExpenseID As String = TextBox1.Text
                Dim Amount As String = TextBox2.Text
                Dim TotalIncome As String = TextBox3.Text
                Dim Description As String = TextBox6.Text
                Dim Tags As String = TextBox4.Text
                Dim Currency As String = ComboBox2.SelectedItem.ToString
                Dim Category As String = TextBox5.Text
                Dim Paymentmethod As String = ComboBox1.SelectedItem.ToString
                Dim Frequency As String = ComboBox5.SelectedItem.ToString()
                Dim ApprovalStatus As String = ComboBox4.SelectedItem.ToString()
                'Dim Receiver As String = ComboBox5.SelectedItem.ToString()
                Dim DateOfexpenses As String = DateTimePicker1.Value


                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Create an OleDbCommand to update the Expense data in the database  
                Dim cmd As New OleDbCommand("UPDATE [Expense] SET [Amount] = ?, [TotalIncome] = ?, [Description] = ?, [Tags] = ?, [Currency] =?, [Category] = ?, [Paymentmethod] = ?, [Frequency] = ?, [ApprovalStatus] = ?, [DateOfexpenses] = ? WHERE [ID] = ?", connect)

                ' Set the parameter values from the UI controls  

                'cmd.Parameters.AddWithValue("@ExpenseID", ExpenseID)
                cmd.Parameters.AddWithValue("@Amount", Amount)
                cmd.Parameters.AddWithValue("@TotalIncome", TotalIncome)
                cmd.Parameters.AddWithValue("@Description", Description)
                cmd.Parameters.AddWithValue("@Tags", Tags)
                cmd.Parameters.AddWithValue("@Currency", Currency)
                cmd.Parameters.AddWithValue("@Category", Category)
                cmd.Parameters.AddWithValue("@PaymentMethod", Paymentmethod)
                cmd.Parameters.AddWithValue("Frequency", Frequency)
                cmd.Parameters.AddWithValue("ApprovalStatus", ApprovalStatus)
                'cmd.Parameters.AddWithValue("Receiver", Receiver)
                cmd.Parameters.AddWithValue("DateOfexpense", DateOfexpenses)
                cmd.Parameters.AddWithValue("ID", ID)


                cmd.ExecuteNonQuery()

                MsgBox("Expense information updated!")

                'Module1.ClearControls(Me)
                'LoadExpenseDataFromDatabase()

            End Using
        Catch ex As OleDbException

            'MessageBox.Show($"Error updating Expense in database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            MessageBox.Show("Error saving Expense to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As FormatException
            Debug.WriteLine($"Format error in btnEdit:{ex.Message}")
            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            MessageBox.Show("Ensure all feilds are filled correctly.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Debug.WriteLine("Failed entering btnEdit_")
            MessageBox.Show("Error saving expense to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            Debug.WriteLine($"An  Error has occured when Editing data from Database")
        End Try
        'LoadExpenseDataFromDatabase()
        Debug.WriteLine("Exited btnEdit")
    End Sub
End Class
