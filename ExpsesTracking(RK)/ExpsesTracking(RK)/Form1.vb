Imports System.Windows.Forms
Imports System.IO
Imports System.Net.Mail
Imports System.Net
Imports System.Data.OleDb

Public Class Form1

    Private mealPlanData As DataTable

    ' Create a ToolTip object
    Private toolTip As New ToolTip()
    Private toolTip1 As New ToolTip()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Initialize ToolTip properties (optional)
        toolTip.AutoPopDelay = 5000
        toolTip.InitialDelay = 500
        toolTip.ReshowDelay = 200
        toolTip.ShowAlways = True

        toolTip1.SetToolTip(Button1, "Cancel")
        toolTip1.SetToolTip(Button2, "Edit")
        toolTip1.SetToolTip(Button3, "Delete")
        toolTip1.SetToolTip(Button4, "Calculate")
        toolTip1.SetToolTip(Button8, "Sort")
        toolTip1.SetToolTip(Button6, "Print as Document")
        toolTip1.SetToolTip(Button7, "Save")


    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Visible = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
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

                Dim ExpenseID As String = TextBox1.Text
                Dim Amount As String = TextBox2.Text
                Dim TotalBudget As String = TextBox3.Text
                Dim Description As String = TextBox4.Text
                Dim Tags As String = TextBox5.Text
                Dim Currency As String = ComboBox1.SelectedItem.ToString
                Dim Category As String = TextBox6.Text
                Dim PaymentMethod As String = ComboBox2.SelectedItem.ToString
                Dim Frequency As String = ComboBox3.SelectedItem.ToString()
                Dim ApprovalStatus As String = ComboBox4.SelectedItem.ToString()
                Dim Receiver As String = ComboBox5.SelectedItem.ToString()
                Dim DateOfExpense As String = DateTimePicker1.Value


                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Create an OleDbCommand to update the Expense data in the database  
                Dim cmd As New OleDbCommand("UPDATE [Expense] SET [ExpenseID] = ?, [Amount] = ?, [TotalBudget] = ?, [Description] = ?, [Tags] = ?, [Currency] =?, [Category] = ?, [PaymentMethod] = ?, [Frequency] = ?, [ApprovalStatus] = ?, [Receiver] = ?, [DateOfExpense] = ? WHERE [ID] = ?", connect)

                ' Set the parameter values from the UI controls  

                cmd.Parameters.AddWithValue("@ExpenseID", ExpenseID)
                cmd.Parameters.AddWithValue("@Amount", Amount)
                cmd.Parameters.AddWithValue("@TotalBudget", TotalBudget)
                cmd.Parameters.AddWithValue("@Description", Description)
                cmd.Parameters.AddWithValue("@Tags", Tags)
                cmd.Parameters.AddWithValue("@Currency", Currency)
                cmd.Parameters.AddWithValue("@Category", Category)
                cmd.Parameters.AddWithValue("@PaymentMethod", PaymentMethod)
                cmd.Parameters.AddWithValue("Frequency", Frequency)
                cmd.Parameters.AddWithValue("ApprovalStatus", ApprovalStatus)
                cmd.Parameters.AddWithValue("Receiver", Receiver)
                cmd.Parameters.AddWithValue("DateOfExpense", DateOfExpense)
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

        Debug.WriteLine("Exited btnEdit")
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
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
            Using connect As New OleDbConnection(connectionString)
                connect.Open()

                ' Update the table name if necessary  
                Dim tableName As String = "Expense"

                ' Create an OleDbCommand to insert the Expense data into the database 
                Dim cmd As New OleDbCommand("INSERT INTO [Expense] ([ExpenseID], [Amount], [TotalBudget], [Description], [Tags], [Currency], [Category], [PaymentMethod], [Frequency], [ApprovalStatus], [Receiver], [DateOfExpense]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", connect)

                ' Set the parameter values from the UI controls 
                'Class declaretions

                Dim expense As New expenses With {
                    .ExpenseID = TextBox1.Text,
                    .Amount = TextBox2.Text,
                    .TotalBudget = TextBox3.Text,
                    .Description = TextBox4.Text,
                    .Tags = TextBox5.Text,
                    .Currency = ComboBox1.SelectedItem.ToString,
                    .Category = TextBox6.Text,
                    .PaymentMethod = ComboBox2.SelectedItem.ToString,
                .Frequency = ComboBox3.SelectedItem.ToString(),
                    .ApprovalStatus = ComboBox4.SelectedItem.ToString(),
                    .Receiver = ComboBox5.SelectedItem.ToString(),
                    .DateOfExpense = DateTimePicker1.Value}

                'txtRecentUpdate.Text = $" Expense updated at {DateTime.Now:HH:MM}"

                cmd.Parameters.Clear()


                cmd.Parameters.AddWithValue("@ExpenseID", expense.ExpenseID)
                cmd.Parameters.AddWithValue("@Amount", expense.Amount)
                cmd.Parameters.AddWithValue("@TotalBudget", expense.TotalBudget)
                cmd.Parameters.AddWithValue("@Description", expense.Description)
                cmd.Parameters.AddWithValue("@Tags", expense.Tags)
                cmd.Parameters.AddWithValue("@Currency", expense.Currency)
                cmd.Parameters.AddWithValue("@Category", expense.Category)
                cmd.Parameters.AddWithValue("@PaymentMethod", expense.PaymentMethod)
                cmd.Parameters.AddWithValue("@Frequency", expense.Frequency)
                cmd.Parameters.AddWithValue("@ApprovalStatus", expense.ApprovalStatus)
                cmd.Parameters.AddWithValue("@Receiver", expense.Receiver)
                cmd.Parameters.AddWithValue("@DateOfExpense", expense.DateOfExpense)


                MsgBox("Expense Information Saved!" & vbCrLf &
                        "ExpenseID: " & expense.ExpenseID & vbCrLf &
                        "Amount: " & expense.Amount & vbCrLf &
                        "TotalBudget: " & expense.TotalBudget & vbCrLf &
                        "Description: " & expense.Description & vbCrLf &
                        "Tags: " & expense.Tags & vbCrLf &
                        "Currency: " & expense.Currency & vbCrLf &
                        "Category: " & expense.Category & vbCrLf &
                        "PaymentMethod: " & expense.PaymentMethod & vbCrLf &
                         "Frequency: " & expense.Frequency & vbCrLf &
                         "ApprovalStatus: " & expense.ApprovalStatus & vbCrLf &
                          "Receiver: " & expense.Receiver & vbCrLf &
                          "DateOfExpense: " & expense.DateOfExpense.ToString, vbInformation, "Expense Confirmation")

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
        connect.Close()
        Debug.WriteLine("Exiting btnSubmit")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Debug.WriteLine("Entering btnDelete_")

        ' Check if there are any selected rows in the DataGridView for expenses  
        If DataGridView1.SelectedRows.Count > 0 Then




            ' Get the selected row  
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Assuming there is an "ID" column which is the primary key in the database  
            Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  

            ' Confirm deletion  
            Dim confirmationResult As DialogResult = MessageBox.Show("Are you sure you want to delete this expense?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If confirmationResult = DialogResult.Yes Then
                ' Proceed with deletion  
                Try
                    TextBox9.Text = $" Expense updated at {DateTime.Now:HH:MM}"

                    Using connect As New OleDbConnection(connectionString)
                        connect.Open()
                        Debug.WriteLine("User confirmed deletion")
                        ' Create the delete command  
                        Dim cmd As New OleDbCommand("DELETE FROM [Expense] WHERE [ID] = ?", connect)
                        cmd.Parameters.AddWithValue("@ID", ID) ' Primary key for matching record  

                        ' Execute the delete command  
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show("Expense deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ' Optionally refresh DataGridView or reload from database  

                            'LoadExpenseDataFromDatabase()

                            TextBox9.Text = $" Expense Delete at {DateTime.Now:HH:MM}"

                        Else
                            Debug.WriteLine(" User canceled deletion")
                            MessageBox.Show("No expense was deleted. Please check if the ID exists.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End Using

                Catch ex As FormatException
                    Debug.WriteLine($"Format error in btnDelete: {ex.Message}")
                    Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
                    MessageBox.Show("Ensure you have selected a row correctly")
                Catch ex As Exception
                    MessageBox.Show($"An error occurred while deleting the expense: {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Debug.WriteLine($"An error has occured when deleteing from database")
                    Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
                End Try
            End If
        Else
            Debug.WriteLine($" No row  selected, exiting btnDelete")
            MessageBox.Show("Please select an expense to delete.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        Debug.WriteLine("Exiting deletion")

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Debug.WriteLine("Entering btnCalculate")

        ' Calculate total expenses  
        Dim totalExpenses As Decimal = Decimal.Parse(TextBox2.Text)
        Dim averageExpenses As Decimal = 0

        Try
            Debug.WriteLine("User confirmed btnCalculate")
            ' Calculate average expenses based on frequency  

            If ComboBox3.SelectedItem IsNot Nothing Then

                Select Case ComboBox3.SelectedItem.ToString()
                    Case "Monthly"
                        averageExpenses = totalExpenses / 12
                    Case "Weekly"
                        averageExpenses = totalExpenses / 52
                    Case "Daily"
                        averageExpenses = totalExpenses / 365
                    Case Else
                        ' Handle other frequencies as needed  
                End Select


            End If

            ' Display total and average expenses on the form  
            Label16.Text = $" R {totalExpenses:N2}"
            Label14.Text = $" {ComboBox3.SelectedItem}: R {averageExpenses:N2}"

        Catch ex As FormatException
            Debug.WriteLine("Invalid format in Button6_Click: Amount should be in numbers.")
            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            MessageBox.Show("Please enter a valid numeriv value for the Amount.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        Catch ex As Exception
            Debug.WriteLine("Failed to calculate")
            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            Debug.WriteLine($"Unexpected error in Button6_Click: {ex.Message}")
            MessageBox.Show("An unexpected error occured during calculations.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Debug.WriteLine($"Calculation complete. Total:{totalExpenses},Avarage:{averageExpenses}")
        Debug.WriteLine("Exiting btnCalculate")
    End Sub



    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        PrintDialog1.Document = PrintDocument1
        If PrintDialog1.ShowDialog() = DialogResult.OK Then
            'LoadFilteredMealPlan() ' Load filtered data based on selected frequency
            If mealPlanData.Rows.Count > 0 Then
                PrintDocument1.Print()
            Else
                MessageBox.Show("No meal plans found for the selected period.", "Print Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        DataGridView1.Sort(DataGridView1.Columns("Amount"), System.ComponentModel.ListSortDirection.Descending)
        DataGridView1.Sort(DataGridView1.Columns("DateOfExpense"), System.ComponentModel.ListSortDirection.Ascending)
    End Sub
End Class
Public Class expenses

    Public Property ExpenseID As Integer
    Public Property Amount As Integer
    Public Property TotalBudget As Integer
    Public Property Description As String
    Public Property Tags As String
    Public Property Currency As String
    Public Property Category As String
    Public Property PaymentMethod As String
    Public Property Frequency As String
    Public Property ApprovalStatus As String
    Public Property Receiver As String
    Public Property DateOfExpense As DateTime

End Class
