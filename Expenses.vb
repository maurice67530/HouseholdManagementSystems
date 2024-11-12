Imports System.IO
Imports System.Windows.Forms
Imports System.Data.OleDb
Public Class Expenses
    Dim conn As New OleDb.OleDbConnection(Module1.connectionString)
    Dim personals As Object
    Dim expensed As Object
    Dim exp As Object

    Private Property expense As Object
    Private Property lblAverageExpenses As Object
    Dim expessed As Object

    Private Property Expensess As Object

    Private Property ID As Object

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs)
        Chores.ShowDialog()

    End Sub

    Private Sub Button1_Click_1(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        DashBoard.Show()
        Me.Close()
    End Sub

    Private Sub btnSubmit_Click(sender As System.Object, e As System.EventArgs) Handles btnSubmit.Click

        ''Create Instance of Expenses Information
        'Dim expense As New Expenses

        'expense.Dates = DateTimePicker1.Value.ToShortDateString 'Get the actual date value
        'expense.Amount = TextBox2.Text
        'expense.Currency = ComboBox1.Text
        'expense.frequency = ComboBox2.Text
        'expense.approvalStatus = ComboBox3.Text
        'expense.Tags = TextBox2.Text
        'expense.payment = ComboBox4.Text




        '' Display a confirmation messageBox  
        'MsgBox("Expense Information Added!" & vbCrLf &
        '        "Amount: " & expense.Amount.ToString & vbCrLf &
        '        "Currency: " & expense.Currency & vbCrLf &
        '        "Payment: " & expense.payment & vbCrLf &
        '        "Tags: " & expense.Tags & vbCrLf &
        '        "Frequency: " & expense.frequency & vbCrLf &
        '        "Approval Status: " & expense.approvalStatus & vbCrLf &
        '        "Dates: " & expense.Dates.ToShortDateString(), vbInformation, "Expense Confirmation")

        'populatedDatagridview()
        ''Module1.LoadPersonDataFromFile("C:\Users\user\Desktop\My Project\My hello\HOUSEHOLD MANAGEMENT SYSTEM.vb1\HOUSEHOLD MANAGEMENT SYSTEM.vb1\Austin.accdb")
        ''Try
        'Using conn As New OleDbConnection(Module1.connectionString)
        '    conn.Open()

        '    Dim tableName As String = "Expenses"

        '    ' Create an OleDbCommand to insert the Expenses data into the database  
        '    Dim cmd As New OleDbCommand("INSERT INTO Expensesdetails{tablename} [ExpensesID], [Dates], [amount], [Currency], [Frequency], [ApprovalStatus],[Tags],[paymentmethod],[person]) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", conn)

        '    ' Set the parameter values from the UI controls

        '    For Each Expensess As Expenses In Expenses
        '        cmd.Parameters.Clear()
        '        cmd.Parameters.AddWithValue("@ExpensesID", Expensess.ID)
        '        cmd.Parameters.AddWithValue("@Dates", Expensess.Dates)
        '        cmd.Parameters.AddWithValue("@amount", Expensess.Amount)
        '        cmd.Parameters.AddWithValue("@currency", Expensess.Currency)
        '        cmd.Parameters.AddWithValue("@Frequency", Expensess.frequency)
        '        cmd.Parameters.AddWithValue("@ApprovalStatus", Expensess.approvalStatus)
        '        cmd.Parameters.AddWithValue("@paymentmethod", Expensess.payment)
        '        cmd.Parameters.AddWithValue("@Tags", Expensess.Tags)



        '        ' Execute the SQL command to insert the data  
        '        cmd.ExecuteNonQuery()

        '    Next
        'End Using
        'Try
        'Catch ex As OleDbException
        '    MessageBox.Show("Error saving personnel to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'Catch ex As Exception
        '    MessageBox.Show("Unexpected Error: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End Try
        'conn.Close()



        '    'Create Instance of Expenses Information  
        '    Dim Expenses As New Expenses()

        '    ' Assign values from textBoxes to the ExpenseInfomation  
        '    expensesID = (TextBox1.Text)
        '    Expenses.Dates = (DateTimePicker1.Value.ToShortDateString())
        '    Expenses.amount = (TextBox2.Text)
        '    Expenses.Currency = (ComboBox1.Text)
        '    Expenses.Frequency = (ComboBox2.Text)
        '    Expenses.ApprovalStatus = (ComboBox2.Text)
        '    Expenses.Tags = (TextBox3.Text)
        '    Expenses.paymentmethod = (ComboBox4.Text)
        '    Expenses.person = (ComboBox5.Text)

        '    'Add the expenses to the List  
        '    'expense.payment = ComboBox4.Text
        '    'expense.Add(Expenses)


        '    ' Display a confirmation messageBox  
        '    MsgBox("Expenses Information added!" & vbCrLf & "ExpensesID:" & Expenses.expensesID & vbCrLf &
        '    "Dates:" & Expenses.Dates & vbCrLf &
        '    "amount:" & Expenses.amount & vbCrLf &
        '    "Currency:" & Expenses.Currency & vbCrLf &
        '    "Frequency:" & Expenses.Frequency & vbCrLf &
        '    "ApprovalStatus:" & Expenses.ApprovalStatus & vbCrLf &
        '    "Tags:" & Expenses.Tags & vbCrLf &
        '    "paymentMethod:" & Expenses.paymentmethod & vbCrLf &
        '    "person:" & Expenses.person, vbInformation, "Expenses confirmation")


        '    'populatedDatagridview()
        '    'Module1.LoadPersonDataFromFile("C:\Users\user\Desktop\My Project\My hello\HOUSEHOLD MANAGEMENT SYSTEM.vb1\HOUSEHOLD MANAGEMENT SYSTEM.vb1\Austin.accdb")
        '    'Try
        '    Using conn As New OleDbConnection(Module1.connectionString)
        '        conn.Open()

        '        Dim tableName As String = "Expenses"

        '        ' Create an OleDbCommand to insert the Expenses data into the database  
        '        Dim cmd As New OleDbCommand("INSERT INTO Expensesdetails{tablename} [ExpensesID], [Dates], [amount], [Currency], [Frequency], [ApprovalStatus],[Tags],[paymentmethod],[person]) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", conn)

        '        ' Set the parameter values from the UI controls

        '        'For Each Expensess As Expenses In expense



        '        cmd.Parameters.Clear()
        '        cmd.Parameters.AddWithValue("@ExpensesID", Expensess.ID)
        '        cmd.Parameters.AddWithValue("@Dates", Expensess.Dates)
        '        cmd.Parameters.AddWithValue("@amount", Expensess.amount)
        '        cmd.Parameters.AddWithValue("@currency", Expensess.Currency)
        '        cmd.Parameters.AddWithValue("@Frequency", Expensess.Frequency)
        '        cmd.Parameters.AddWithValue("@ApprovalStatus", Expensess.ApprovalStatus)
        '        cmd.Parameters.AddWithValue("@paymentmethod", Expensess.paymentmethod)
        '        cmd.Parameters.AddWithValue("@Tags", Expensess.Tags)
        '        cmd.Parameters.AddWithValue("@person", Expensess.person)

        '        ' Execute the SQL command to insert the data  
        '        cmd.ExecuteNonQuery()

        '        'Next
        '    End Using
        '    Try
        '    Catch ex As OleDbException
        '        MessageBox.Show("Error saving personnel to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    Catch ex As Exception
        '        MessageBox.Show("Unexpected Error: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    En
        'End Sub
        'update the expense details with values from controls





        Try

            Dim expense As New Expenses() With {
 .Amount = TextBox2.Text,
 .Currency = ComboBox2.Text,
 .payment = ComboBox1.Text,
 .Tags = TextBox3.Text,
 .frequency = ComboBox3.Text,
 .ApprovalStatus = ComboBox4.Text,
 .Dates = DateTimePicker1.Value}



            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()
                Dim tablename As String = "Expenses"

                Dim Cmd As New OleDbCommand($"INSERT INTO {tablename} ([Amount], [Currency], [Payment], [Tags], [Frequency], [ApprovalStatus], [Dates])VALUES (?, ?, ?, ?, ?, ?, ?)", conn)

                Cmd.Parameters.Clear()

                Cmd.Parameters.AddWithValue("@Amount", TextBox2.Text)
                Cmd.Parameters.AddWithValue("@Currency", ComboBox2.Text())
                Cmd.Parameters.AddWithValue("@Payment", ComboBox1.Text())
                Cmd.Parameters.AddWithValue("@Tags", TextBox3.Text)
                Cmd.Parameters.AddWithValue("@Frequency", ComboBox3.Text())
                Cmd.Parameters.AddWithValue("@ApprovalStatus", ComboBox4.Text())
                Cmd.Parameters.AddWithValue("@Dates", DateTimePicker1.Value)

                Cmd.ExecuteNonQuery()

                'MsgBox("Expense Information Added!" & vbCrLf &
                '"Amount: " & Integer.Parse(TextBox2.Text).ToString & vbCrLf &
                '"Currency: " & expense.Currency & vbCrLf &
                '"Payment: " & expense.payment & vbCrLf &
                '"Tags: " & expense.Tags & vbCrLf &
                '"Frequency: " & expense.frequency & vbCrLf &
                '"Approval Status: " & expense.approvalStatus & vbCrLf &
                '"Date: " & DateTimePicker1.Value.ToShortDateString(), vbInformation, "Expense Confirmation")


                'MessageBox.Show("Expense saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

                ''    End Using
                ''Catch expense As OleDbException
                ''    MessageBox.show($"Error Saving to Database: {expense.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

                ''    'Catch expense As OleDbException
                '    MessageBox.Show($"Unexpected error:{expense.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                'End Try
            End Using
        Catch ex As OleDbException
            Debug.WriteLine($"Database error in btnSubmit_Click:{ex.Message}")
            MessageBox.Show("Error saving personnel to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Debug.WriteLine($"general error in btnsubmit_click:{ex.Message}")
            MessageBox.Show("Unexpected Error: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        conn.Close()
    End Sub

    Public Sub loadExpensesdataFromDatabase()
        Try
            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()

                Dim tableName As String = "Expenses"

                Dim cmd As New OleDbCommand($"SELECT * FROM {tableName}", conn)

                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)

                DataGridView1.DataSource = dt

            End Using

        Catch ex As Exception
            MessageBox.Show($"Error Loading  Expenses data from database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)


        End Try
    End Sub


    Private Sub populatedDatagridview()


        ' Clear Existing rows   
        DataGridView1.Rows.Clear()

        ' Add each expense to the DataGridView  
        For Each exp As Expenses In Expenses

            DataGridView1.Rows.Add(exp.Amount, exp.Currency, exp.payment.ToString, exp.Tags, exp.frequency, exp.ApprovalStatus, exp.Tags, exp.Dates())
        Next
    End Sub
    'Private Expenses As New List(Of Expenses)
    Private Sub PopulateComboBox()
        ComboBox2.Items.Clear()
        For Each Expenses As Expenses In expense
            ComboBox2.Items.Add(expense.FirstName & " " & expense.Lastname)
        Next
    End Sub


    Private Expenses As New List(Of Expenses)

    Private Sub Expenses_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        'Module1.LoadPersonDataFromFile("C:\Users\user\Documents\Visual Studio 2010\Projects\My hello\HOUSEHOLD MANAGEMENT SYSTEM.vb1\HOUSEHOLD MANAGEMENT SYSTEM.vb1\Storage\Person.txt")
        loadExpensesdataFromDatabase()
        'PopulateComboBox()

    End Sub

    Private Sub btnGetINput_Click(sender As System.Object, e As System.EventArgs) Handles btnGetINput.Click

        TextBox3.Text = InputBox("Enter your amount")
        TextBox2.Text = InputBox("Enter Your Tags")

    End Sub

    'Private Sub LoadExpensesDataFromFile(filePath As String)
    '    If File.Exists(filePath) Then
    '        Try
    '            Using reader As New StreamReader(filePath)
    '                Dim Expenses As Expenses = Nothing
    '                While Not reader.EndOfStream
    '                    Dim line As String = reader.ReadLine()
    '                    If line.StartsWith("Name: ") Then
    '                        If Expenses IsNot Nothing Then
    '                            expense.Add(Expenses) ' Add the previous person to the list first  
    '                        End If
    '                        expense = New Expenses()
    '                        Expenses.Dates = line.Substring("Dates: ".Length).Trim()
    '                    ElseIf line.StartsWith("amount: ") Then
    '                        Expenses.amount = line.Substring("amount: ".Length).Trim()
    '                    ElseIf line.StartsWith("Currency ") Then
    '                        Expenses.Frequency = Integer.Parse(line.Substring("Frequency: ".Length).Trim())
    '                    ElseIf line.StartsWith("ApprovalStatus: ") Then
    '                        Expenses.ApprovalStatus = line.Substring("Approvalstatus: ".Length).Trim()
    '                    ElseIf line.StartsWith("tags: ") Then
    '                        Expenses.Tags = line.Substring("Tags: ".Length).Trim()
    '                    ElseIf line.StartsWith("paymentmethod: ") Then
    '                        Expenses.payment = line.Substring("paymentmethod: ".Length).Trim()

    '                    End If
    '                End While

    '                ' Add the last person to the list if we reached the end of the file  

    '                If person IsNot Nothing Then
    '                    expense.Add(Expenses)
    '                End If
    '            End Using
    '        Catch ex As Exception
    '            MessageBox.Show("Error loading personnel data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        End Try
    '    Else
    '        MessageBox.Show("Personnel file Not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End If
    'End Sub

    Private Sub btnCalculate_Click(sender As System.Object, e As System.EventArgs) Handles btnCalculate.Click

        'If Expenses.Count > 0 Then
        '    ' Calculate total expenses
        '    Dim totalExpenses As Integer = Expenses.Sum(Function(exp) Convert.ToInt32(exp.amount))

        '    ' Calculate average expenses
        '    Dim averageExpenses As Double = Expenses.Average(Function(exp) Convert.ToDouble(exp.amount))

        '    ' Display results in labels
        '    Label11.Text = "Total Expenses: $" & totalExpenses.ToString()
        '    Label12.Text = "Average Expense: $" & averageExpenses.ToString("F2")
        'Else
        '    MsgBox("No expenses available for calculation.")
        'End If


        'Calculate total expenses  
        Dim totalExpenses As Decimal = Decimal.Parse(TextBox2.Text)

        ' Calculate average expenses based on frequency  
        Dim averageExpenses As Decimal = 0
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
        Label11.Text = $"Total Expense: R {totalExpenses:N2}"
        Label12.Text = $"Average Expense per {ComboBox3.SelectedItem}: R {averageExpenses:N2}"

        Debug.WriteLine("Entering btnCalculate_Click")
        'Existing calculation code...
        Debug.WriteLine($"calculation comlete Total:{totalExpenses}, Avarage:{averageExpenses}")
        Debug.WriteLine("Exiting btnCalculate_click")



        Try

        Catch ex As FormatException
            Debug.WriteLine("Invalid format in btnCalculate_click:Amount should be a number")
            MessageBox.Show("please enter a  valid numeric number for the amount.", "Input error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Catch ex As Exception
            Debug.WriteLine($"Unexpected error in btnCalculate_click:{ex.Message}")
            MessageBox.Show("An unexpected error occured during calculation,", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub btnClear_Click(sender As System.Object, e As System.EventArgs) Handles btnClear.Click
        TextBox2.Clear()
        TextBox3.Clear()
    End Sub


    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click

        ' Check if there are any selected rows in the DataGridView for expenses  
        If DataGridView1.SelectedRows.Count > 0 Then
            ' Get the selected row  
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Assuming there is an "ID" column which is the primary key in the database  
            Dim expenseId As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  

            ' Confirm deletion  
            Dim confirmationResult As DialogResult = MessageBox.Show("Are you sure you want to delete this expense?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If confirmationResult = DialogResult.Yes Then
                ' Proceed with deletion  
                Try
                    Using conn As New OleDbConnection(Module1.connectionString)
                        conn.Open()

                        ' Create the delete command  
                        Dim cmd As New OleDbCommand("DELETE FROM [Expenses] WHERE [ID] = ?", conn)
                        cmd.Parameters.AddWithValue("@ID", expenseId) ' Primary key for matching record  

                        ' Execute the delete command  
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show("Expense deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ' Optionally refresh DataGridView or reload from database  
                            'PopulateDataGridView()
                        Else
                            MessageBox.Show("No expense was deleted. Please check if the ID exists.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End Using

                Catch ex As Exception
                    MessageBox.Show($"An error occurred while deleting the expense: {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            MessageBox.Show("Please select an expense to delete.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            'load the data from  the selected row into UI controls 

            TextBox2.Text = selectedRow.Cells("Amount").Value.ToString()
            ComboBox1.SelectedItem = selectedRow.Cells("currency").Value.ToString()
            ComboBox2.SelectedItem = selectedRow.Cells("payment").Value.ToString()
            TextBox3.Text = selectedRow.Cells("Tags").Value.ToString()
            ComboBox3.SelectedItem = selectedRow.Cells("Frequency").Value.ToString()
            ComboBox4.SelectedItem = selectedRow.Cells("ApprovalStatus").Value.ToString()
            DateTimePicker1.Text = selectedRow.Cells("Dates").Value.ToString()

        End If
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click

        If DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Dim Amount As String =TextBox2.Text ' Convert amount to Integer  
            Dim Currency As String = ComboBox1.Text
            Dim payment As String = ComboBox2.Text
            Dim Tags As String = TextBox3.Text
            Dim Frequency As String = ComboBox3.Text
            Dim approvalStatus As String = ComboBox4.Text
            Dim Dates As DateTime = DateTimePicker1.Text ' Ensure this is of DateTime type  

            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()

                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Create an OleDbCommand to update the expense data in the database  
                Dim cmd As New OleDbCommand("UPDATE [Expenses] SET [Amount] = ?, [Currency]  = ?, [payment] = ?, [Tags] = ?, [Frequency] = ?, [ApprovalStatus] = ?, [Dates] = ? WHERE [ID] = ?", conn)

                ' Set the parameter values from the UI controls  

                cmd.Parameters.AddWithValue("@Amount", TextBox2.Text)
                cmd.Parameters.AddWithValue("@Currency", ComboBox1.Text)
                cmd.Parameters.AddWithValue("@payment", ComboBox2.SelectedItem.ToString())
                cmd.Parameters.AddWithValue("@Tags", TextBox3.Text)
                cmd.Parameters.AddWithValue("@Frequency", ComboBox3.SelectedItem.ToString())
                cmd.Parameters.AddWithValue("@Approvalstatus", ComboBox4.SelectedItem.ToString())
                cmd.Parameters.AddWithValue("@Dates", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@ID", ID)

                'cmd.Parameters.AddWithValue("@ID", ExpensesID) ' Primary key for matching record  
                cmd.ExecuteNonQuery()


                MsgBox("Expense Updated Successfuly!", vbInformation, "Update Confirmation")

                loadExpensesdataFromDatabase()
                expense.ClearControls(Me)

            End Using
        Catch ex As OleDbException
            MessageBox.Show($"Error updating Expenses in database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs)

    End Sub
End Class
