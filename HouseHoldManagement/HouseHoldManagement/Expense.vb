﻿Imports System.IO
Imports System.Net.Mail
Imports System.Net
Imports System.Data.OleDb
Public Class Expense
    'Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\khodani\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"
    Dim conn As New OleDbConnection(Rinae.connectionString)

    ' Create a ToolTip object
    Private toolTip As New ToolTip()
    Private toolTip1 As New ToolTip()

    Private mealPlanData As DataTable

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
            Using conn As New OleDbConnection(Rinae.connectionString)
                conn.Open()

                ' Update the table name if necessary  
                Dim tableName As String = "Expense"

                ' Create an OleDbCommand to insert the Expense data into the database 
                Dim cmd As New OleDbCommand("INSERT INTO [Expense] ([Amount], [TotalIncome], [Description], [Tags], [Currency], [Category], [Paymentmethod], [Frequency], [ApprovalStatus], [DateOfexpenses], {Person}) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", conn)

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
                    .Person = ComboBox3.SelectedItem.ToString(),
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
                cmd.Parameters.AddWithValue("@PaymentMethod", expense.Paymentmethod)
                cmd.Parameters.AddWithValue("@Frequency", expense.Frequency)
                cmd.Parameters.AddWithValue("@ApprovalStatus", expense.ApprovalStatus)
                cmd.Parameters.AddWithValue("@Person ", expense.Person)
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
                          "Receiver: " & expense.Person & vbCrLf &
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
        LoadExpenseDataFromDatabase()
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
            Using connect As New OleDbConnection(Rinae.connectionString)
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
                Dim Person As String = ComboBox3.SelectedItem.ToString()
                Dim DateOfexpenses As String = DateTimePicker1.Value


                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Create an OleDbCommand to update the Expense data in the database  
                Dim cmd As New OleDbCommand("UPDATE [Expense] SET [Amount] = ?, [TotalIncome] = ?, [Description] = ?, [Tags] = ?, [Currency] =?, [Category] = ?, [Paymentmethod] = ?, [Frequency] = ?, [ApprovalStatus] = ?, [DateOfexpenses] = ?, [Person] = ? WHERE [ID] = ?", connect)

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
                cmd.Parameters.AddWithValue("Person ", Person)
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
        LoadExpenseDataFromDatabase()
        Debug.WriteLine("Exited btnEdit")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
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
                    TextBox7.Text = $" Expense updated at {DateTime.Now:HH:MM}"

                    Using connect As New OleDbConnection(Rinae.connectionString)
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

                            TextBox7.Text = $" Expense Delete at {DateTime.Now:HH:MM}"

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
        LoadExpenseDataFromDatabase()

        Debug.WriteLine("Exiting deletion")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        DataGridView1.Sort(DataGridView1.Columns("Amount"), System.ComponentModel.ListSortDirection.Descending)
        DataGridView1.Sort(DataGridView1.Columns("Currency"), System.ComponentModel.ListSortDirection.Ascending)

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox6.Text = ""
        TextBox4.Text = ""
        ComboBox2.SelectedItem = ""
        TextBox5.Text = ""
        ComboBox1.SelectedItem = ""
        ComboBox5.SelectedItem = ""
        ComboBox4.SelectedItem = ""
        TextBox7.Text = ""
        ComboBox3.SelectedItem = ""

    End Sub

    Public Sub LoadExpenseDataFromDatabase()

        Debug.WriteLine("LoadMealPlansDataFromDatabase")
        Using connect As New OleDbConnection(Rinae.connectionString)
            connect.Open()

            ' Update the table name if necessary  
            Dim tableName As String = "Expense"

            ' Create an OleDbCommand to select the data from the database  
            Dim cmd As New OleDbCommand($"SELECT * FROM {tableName}", connect)

            ' Create a DataAdapter and fill a DataTable  
            Dim da As New OleDbDataAdapter(cmd)
            Dim dt As New DataTable()
            da.Fill(dt)

            ' Bind the DataTable to the DataGridView  
            DataGridView1.DataSource = dt
            'HighlightExpiredItemss()
        End Using

    End Sub
    Private Sub Expense_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Initialize ToolTip properties (optional)
        toolTip.AutoPopDelay = 5000
        toolTip.InitialDelay = 500
        toolTip.ReshowDelay = 200
        toolTip.ShowAlways = True

        toolTip1.SetToolTip(Button5, "Sort")
        toolTip1.SetToolTip(Button6, "Clear controls")
        toolTip1.SetToolTip(Button3, "Edit")
        toolTip1.SetToolTip(Button4, "Delete")
        toolTip1.SetToolTip(Button7, "Calculate Budget")
        'toolTip1.SetToolTip(Button7, "Daily tasks")
        toolTip1.SetToolTip(Button1, "Save")
        LoadExpenseDataFromDatabase()
        PopulateComboboxFromDatabase(ComboBox3)
    End Sub

    Public Sub PopulateComboboxFromDatabase(ByRef comboBox As ComboBox)
        Dim conn As New OleDbConnection(Rinae.connectionString)
        Try
            Debug.WriteLine("populate combobox successful")
            'open the database connection
            conn.Open()

            'retrieve the firstname and surname columns from the personaldetails tabel
            Dim query As String = "SELECT FirstName, LastName FROM PersonalDetails"
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            'bind the retrieved data to the combobox
            ComboBox3.Items.Clear()
            While reader.Read()
                ComboBox3.Items.Add($"{reader("FirstName")} {reader("LastName")}")
            End While

            'close the database
            reader.Close()

        Catch ex As Exception
            'handle any exeptions that may occur  
            Debug.WriteLine("failed to populate combobox")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error: {ex.StackTrace}")

        Finally
            'close the database connection
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Debug.WriteLine("Entering btnCalculate")

        ' Calculate total expenses  
        Dim totalExpenses As Decimal = Decimal.Parse(TextBox2.Text)
        Dim averageExpenses As Decimal = 0

        Try
            Debug.WriteLine("User confirmed btnCalculate")
            ' Calculate average expenses based on frequency  

            If ComboBox5.SelectedItem IsNot Nothing Then

                Select Case ComboBox5.SelectedItem.ToString()
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
            Label15.Text = $"Total Expense: R {totalExpenses:N2}"
            'Label16.Text = $" {ComboBox5.SelectedItem}: R {averageExpenses:N2}"

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

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
 Try

            Debug.WriteLine("selecting data in the datagridview")
            If DataGridView1.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

                ' Load the data from the selected row into UI controls  
                TextBox2.Text = selectedRow.Cells("Amount").Value.ToString()
                TextBox3.Text = selectedRow.Cells("TotalIncome").Value.ToString()
                TextBox6.Text = selectedRow.Cells("Description").Value.ToString()
                TextBox4.Text = selectedRow.Cells("Tags").Value.ToString()
                ComboBox2.SelectedItem = selectedRow.Cells("currency").Value.ToString()
                TextBox5.Text = selectedRow.Cells("Category").Value.ToString()
                ComboBox1.SelectedItem = selectedRow.Cells("Paymentmethod").Value.ToString()
                ComboBox5.SelectedItem = selectedRow.Cells("Frequency").Value.ToString()
                ComboBox4.SelectedItem = selectedRow.Cells("Approvalstatus").Value.ToString()
                DateTimePicker1.Text = selectedRow.Cells("Dateofexpenses").Value.ToString()
                ComboBox3.SelectedItem = selectedRow.Cells("Person").Value.ToString()
            End If
        Catch ex As Exception
            Debug.WriteLine("error selection data in the database")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
        End Try
    End Sub
    Public Sub ClearControlss(ByVal FORM As Form)
        ' Clear TextBoxes  
        For Each ctrl As Control In FORM.Controls
            If TypeOf ctrl Is TextBox Then
                CType(ctrl, TextBox).Clear()
            End If
        Next

        ' Clear ComboBoxes  
        For Each ctrl As Control In FORM.Controls
            If TypeOf ctrl Is ComboBox Then
                CType(ctrl, ComboBox).ResetText()
            End If
        Next

        ' Clear DateTimePickers  
        For Each ctrl As Control In FORM.Controls
            If TypeOf ctrl Is DateTimePicker Then
                CType(ctrl, DateTimePicker).Value = DateTimePicker.MinimumDateTime ' or set to a specific date  
            End If
        Next
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        LoadExpenseDataFromDatabase()
    End Sub
End Class

