Imports System.IO
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Net.Mail
Imports System.Net



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
    Public Property DateAdded As Object
    Public Property Names As String

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs)
        Chores.ShowDialog()

        Try
            Debug.WriteLine("Entering btnSubmit successfuly")

            Dim txtRecentEdit As New TextBox()
            txtRecentEdit.Text = $"Expense saved at {DateTime.Now:HH:MM}"
            txtRecentEdit.Location = New Point(65, Label3.Top + 25)
            txtRecentEdit.AutoSize = False
            txtRecentEdit.Font = New Font("Microsoft Sans serif", 9, FontStyle.Regular)
            txtRecentEdit.ForeColor = Color.Black
            Me.Controls.Add(txtRecentEdit)


            Dim expense As New Expenses() With {
             .ItemID = TextBox2.Text,
             .Currency = ComboBox2.Text,
             .Payment = ComboBox1.Text,
             .Tags = TextBox3.Text,
             .Frequency = ComboBox3.Text,
             .ApprovalStatus = ComboBox4.Text,
             .ReorderLevel = txtRecentEdit.Text,
              .Edit = DateTimePicker1.Value}



            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()
                Dim tablename As String = "Expenses"

                Dim Cmd As New OleDbCommand($"INSERT INTO {tablename} ([Amount], [Currency], [Payment], [Tags], [Frequency], [ApprovalStatus], [Dates], [Edit])VALUES ( ?, ?, ?, ?, ?, ?, ?, ?)", conn)

                Cmd.Parameters.Clear()

                Cmd.Parameters.AddWithValue("@Amount", TextBox2.Text)
                Cmd.Parameters.AddWithValue("@Currency", ComboBox2.Text())
                Cmd.Parameters.AddWithValue("@Payment", ComboBox1.Text())
                Cmd.Parameters.AddWithValue("@Tags", TextBox3.Text)
                Cmd.Parameters.AddWithValue("@Frequency", ComboBox3.Text())
                Cmd.Parameters.AddWithValue("@ApprovalStatus", ComboBox4.Text())
                Cmd.Parameters.AddWithValue("@Dates", DateTimePicker1.Value)
                Cmd.Parameters.AddWithValue("@Edit", expense.ReorderLevel)
                Cmd.ExecuteNonQuery()

                MsgBox("Expense Information Added!" & vbCrLf &
                "Amount: " & Integer.Parse(TextBox2.Text).ToString & vbCrLf &
                "Currency: " & expense.Currency & vbCrLf &
                "Payment: " & expense.Payment & vbCrLf &
                "Tags: " & expense.Tags & vbCrLf &
                "Frequency: " & expense.Frequency & vbCrLf &
                "Approval Status: " & expense.ApprovalStatus & vbCrLf &
                "Date: " & DateTimePicker1.Value.ToShortDateString(), vbInformation, "Expense Confirmation")


                MessageBox.Show("Expense saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

                ''    End Using
                ''Catch expense As OleDbException
                ''    MessageBox.show($"Error Saving to Database: {expense.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)


                '    MessageBox.Show($"Unexpected error:{expense.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                'End Try
            End Using
        Catch ex As OleDbException
            Debug.WriteLine($"Database error in btnSubmit_Click:{ex.Message}")
            Debug.WriteLine($"Stack Trace:{ex.StackTrace}")
            MessageBox.Show("Error saving Expense to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Debug.WriteLine($"general error in btnsubmit_click:{ex.Message}")
            Debug.WriteLine($"stack Trace:{ex.StackTrace}")
            MessageBox.Show("Unexpected Error occured: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Debug.WriteLine("Fail to save")

            MessageBox.Show("Error saving to database;" & ex.Message, "saving to database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine("Exiting btnSubmit successfully")
        End Try

        conn.Close()
        loadExpensesdataFromDatabase()
    End Sub

    Public Sub loadExpensesdataFromDatabase()
        Try
            Debug.WriteLine("populated connected succesfully")
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

            Debug.WriteLine("populated fail to initialize succesfully")
            MessageBox.Show($"Error Loading  Expenses data from database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub populatedDatagridview()

        Try
            Debug.WriteLine("poulating datagridview successfully")
            ' Clear Existing rows   
            DataGridView1.Rows.Clear()

            ' Add each expense to the DataGridView  
            For Each exp As Expenses In Expenses

                DataGridView1.Rows.Add(exp.Amount, exp.Currency, exp.Payment.ToString, exp.Tags, exp.Frequency, exp.ApprovalStatus, exp.Edit, exp.Edit())


            Next
        Catch ex As Exception
            Debug.WriteLine("fail to populate")
            Debug.WriteLine($"Stack Trace:{ex.StackTrace}")
            MessageBox.Show("An unexpexted error occured during DataGridview")

        End Try
    End Sub
    'Private Expenses As New List(Of Expenses)
    Private Sub PopulateComboBox()
        Try

            Debug.WriteLine("Combobox loaded successfully")

            ComboBox5.Items.Clear()
            For Each Expenses As Expenses In expense
                ComboBox5.Items.Add(expense.FirstName & " " & expense.Lastname)
            Next
        Catch ex As Exception
            Debug.WriteLine("fail to populate combobox")
            Debug.WriteLine($"Stack Trace:{ex.StackTrace}")
            MessageBox.Show("An unexpexted error occured during populating combobox.", "ComboBoxLoad", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    Private Expenses As New List(Of Expenses)

    Private Sub Expenses_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        ' Set tooltips For buttons
        ToolTip1.SetToolTip(Button1, "Dashboard")
        ToolTip1.SetToolTip(btnSubmit, "Save")
        ToolTip1.SetToolTip(btnDelete, "Delete")
        ToolTip1.SetToolTip(Button2, "Filter")
        ToolTip1.SetToolTip(btnEdit, "Edit")
        ToolTip1.SetToolTip(btnCalculate, "Calculate")
        ToolTip1.SetToolTip(btnClear, "Clear")
        ToolTip1.SetToolTip(btnGetINput, "Input")
        ToolTip1.SetToolTip(BtnImport, "Import")
        ToolTip1.SetToolTip(BtnExport, "Export")

        Try


            'Dim txtRecentEdit As New TextBox()
            'txtRecentEdit.Text = $"Expense saved at {DateTime.Now:HH:MM}"
            'txtRecentEdit.Location = New Point(65, Label3.Top + 25)
            'txtRecentEdit.AutoSize = False
            'txtRecentEdit.Font = New Font("Microsoft Sans serif", 9, FontStyle.Regular)
            'txtRecentEdit.ForeColor = Color.Black
            'Me.Controls.Add(txtRecentEdit)





            Debug.WriteLine("Form loaded successfully")

            'Dymamic Label creation

            Dim lblDateTime As New TextBox()
            lblDateTime.Name = "lblDateTime"
            lblDateTime.Text = "Today: " & DateTime.Now.ToString("dddd, dd MMMM yyyy HH: ss")
            lblDateTime.Location = New Point(10, 10) 'Adjust position as needed
            lblDateTime.AutoSize = True
            lblDateTime.Font = New Font("Arial", 10, FontStyle.Bold)
            lblDateTime.ForeColor = Color.Blue


            'Add the label to the form's controls
            Me.Controls.Add(lblDateTime)
            Debug.WriteLine("Expense_load: DateTime label added.")

            'Dim lblfeedback As New Label()
            'lblfeedback.Text = $"Expenses update at {DateTime.Now:HH:mm:ss}"
            'lblfeedback.Location = New Point(10, DataGridView1.Bottom + 10)
            'lblfeedback.AutoSize = True
            'lblfeedback.Font = New Font("Arial", 10, FontStyle.Italic)
            'lblfeedback.ForeColor = Color.Green
            'Me.Controls.Add(lblfeedback)


            'Module1.LoadPersonDataFromFile("C:\Users\user\Documents\Visual Studio 2010\Projects\My hello\HOUSEHOLD MANAGEMENT SYSTEM.vb1\HOUSEHOLD MANAGEMENT SYSTEM.vb1\Storage\Person.txt")
            loadExpensesdataFromDatabase()
            'PopulateComboBox()
            'MessageBox.Show("Form load finally initialized")

        Catch ex As Exception
            Debug.WriteLine("Form failed to load successfully")
            Debug.WriteLine($"Error in expense_load when creating  DateTime label:{ex.Message}")
            Debug.WriteLine($"Stack Trace:{ex.StackTrace}")
            MessageBox.Show("An unexpexted error occured during Form load.", "Expensesload Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show("An error occured while displaying the date  and time. ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

        '' Define the low stock threshold
        'Dim lowStockThreshold As Integer = 100

        '' Get low stock items from database
        'Dim lowStockItems As List(Of String) = GetLowStockItems(lowStockThreshold)

        '' Send an email if any items are low on stock
        'If lowStockItems.Count > 0 Then
        '    Dim messageBody As String = "The following items are low in stock:" & vbCrLf & String.Join(vbCrLf, lowStockItems)
        '    SendEmail("dongolamaano3@gmail.com", "Expense Alert: Low Stock", messageBody)
        'Else
        '    MessageBox.Show("Expenses are sufficient.", "Expense Check", MessageBoxButtons.OK, MessageBoxIcon.Information)
        'End If
        Dim expenseThreshold As Decimal = budgetLimit * 0.8D
        Dim expenseDetails As String = GetTotalExpensesWithTags()

        ' Extract the total expenses from the last line of the expenseDetails string.
        Dim totalExpenses As Decimal = 0
        Dim lines() As String = expenseDetails.Split({vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
        For Each line As String In lines
            If line.StartsWith("Total Expenses:") Then
                Dim expenseValueString As String = line.Replace("Total Expenses:", "").Trim()
                Decimal.TryParse(expenseValueString, Globalization.NumberStyles.Currency, Nothing, totalExpenses)
                Exit For
            End If
        Next

        If totalExpenses > expenseThreshold Then
            Dim messageBody As String = $"Alert! Expenses exceeded 80% of the budget.{vbCrLf}" &
                $"Budget Limit: {budgetLimit:C}{vbCrLf}{vbCrLf}" &
                                            $"Expense Breakdown:{vbCrLf}{expenseDetails}"
            SendEmail("austinmulalo113@gmail.com", "Budget Alert: High Expenses", messageBody)
        Else
            MessageBox.Show("Expenses are within budget.", "Budget Check", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If


    End Sub
    ' Declarations

    Private connectionString As String = Module1.connectionString
    Private budgetLimit As Decimal = 600
    Private emailFrom As String = "austinmulalo113@gmail.com"
    Private emailPassword As String = "oqsa qwqa bhjc nzoe" ' Use your Gmail/App Password here
    ' Retrieves expense details grouped by Tags and calculates the total
    Private Function GetTotalExpensesWithTags() As String
        Dim totalExpenses As Decimal = 0
        Dim expenseDetails As New List(Of String)
        Try
            Using conn As New OleDbConnection(Module1.connectionString)

                conn.Open()
                Dim query As String = "SELECT Currency, Amount  FROM Expenses "
                Using cmd As New OleDbCommand(query, conn)
                    Using reader As OleDbDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim tag As String = reader("Currency").ToString()
                            Dim amount As Decimal = Convert.ToDecimal(reader("Amount"))
                            totalExpenses += amount
                            expenseDetails.Add($"{tag}: {amount:C}")

                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Database error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Dim expenseSummary As String = String.Join(vbCrLf, expenseDetails)
        Return $"{expenseSummary}{vbCrLf}Total Expenses: {totalExpenses:C}"
    End Function
    Private Sub SendEmail(toEmail As String, subject As String, body As String)

        Try
            Dim smtpClient As New SmtpClient("smtp.gmail.com") With {
                            .Port = 587,
                          .Credentials = New Net.NetworkCredential("austinmulalo113@gmail.com", "oqsa qwqa bhjc nzoe"), ' Replace with Gmail/App Password
                           .EnableSsl = True}

            Dim mail As New MailMessage() With {
                .From = New MailAddress("austinmulalo113@gmail.com"),
                .Subject = subject,
                .Body = body
            }

            mail.To.Add(toEmail)
            smtpClient.Send(mail)
            MessageBox.Show("Email sent successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Failed to send email: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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

        Try
            Debug.WriteLine("Entering btnCalculate")
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


            'Existing calculation code...
            Debug.WriteLine($"calculation complete Total:{totalExpenses}, Avarage:{averageExpenses}")
            Debug.WriteLine("Exiting btnCalculate")


        Catch ex As FormatException
            Debug.WriteLine("Invalid format in btnCalculate_click:Amount should be a number")
            Debug.WriteLine($"stack Trace:{ex.StackTrace}")
            MessageBox.Show("please enter a  valid numeric number for the amount.", "Input error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Catch ex As Exception
            Debug.WriteLine($"Unexpected error in btnCalculate_click:{ex.Message}")
            Debug.WriteLine($"Stack Trace:{ex.StackTrace}")
            MessageBox.Show("An unexpected error occured during calculation,", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine("Exiting btnCalculate")
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
                    Debug.WriteLine(" Entering Delete initialized succesfully")

                    If DataGridView1.SelectedRows.Count > 0 Then
                        Dim confirmationResults As DialogResult = MessageBox.Show("Are you sure you want to delete this expenses?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                        If confirmationResult = DialogResult.Yes Then
                            Debug.WriteLine("User confirmed deletion")
                            'proceed  with delection logic
                        End If
                    End If


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
                    Debug.WriteLine("Fail to initialize Delete successfully")
                    MessageBox.Show($"An error occurred while deleting the expense: {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            Debug.WriteLine("User canceled deletion.")

            Debug.WriteLine("No Row selected,exiting btnDelete_click")
            MessageBox.Show("Please select an expense to delete.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

        Debug.WriteLine("Exiting btnDelete_click")
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
            TextBox1.Text = selectedRow.Cells("Edit").Value.ToString

        End If
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click

        If DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try

            Debug.WriteLine("Entering BtnEdit")

            Dim txtRecentEdit As New TextBox()
            txtRecentEdit.Text = $"Expense udated at {DateTime.Now:HH:MM}"
            txtRecentEdit.Location = New Point(65, Label3.Top + 25)
            txtRecentEdit.AutoSize = False
            txtRecentEdit.Font = New Font("Microsoft Sans serif", 9, FontStyle.Regular)
            txtRecentEdit.ForeColor = Color.Black
            Me.Controls.Add(txtRecentEdit)


            TextBox1.Text = $"expense updated by Austin at {DateTime.Now:hh:mm:ss}"
            Debug.WriteLine("Edit is succesfully initialized")
            Dim Amount As String = TextBox2.Text ' Convert amount to Integer  
            Dim Currency As String = ComboBox1.Text
            Dim payment As String = ComboBox2.Text
            Dim Tags As String = TextBox3.Text
            Dim Frequency As String = ComboBox3.Text
            Dim approvalStatus As String = ComboBox4.Text
            Dim Dates As DateTime = DateTimePicker1.Text ' Ensure this is of DateTime type  
            Dim Edit As String = txtRecentEdit.Text

            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()

                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Create an OleDbCommand to update the expense data in the database  
                Dim cmd As New OleDbCommand("UPDATE [Expenses] SET [Amount] = ?, [Currency]  = ?, [payment] = ?, [Tags] = ?, [Frequency] = ?, [ApprovalStatus] = ?, [Dates] = ?, [Edit] = ? WHERE [ID] = ?", conn)

                ' Set the parameter values from the UI controls  

                cmd.Parameters.AddWithValue("@Amount", TextBox2.Text)
                cmd.Parameters.AddWithValue("@Currency", ComboBox1.Text)
                cmd.Parameters.AddWithValue("@payment", ComboBox2.SelectedItem.ToString())
                cmd.Parameters.AddWithValue("@Tags", TextBox3.Text)
                cmd.Parameters.AddWithValue("@Frequency", ComboBox3.SelectedItem.ToString())
                cmd.Parameters.AddWithValue("@Approvalstatus", ComboBox4.SelectedItem.ToString())
                cmd.Parameters.AddWithValue("@Dates", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@feedback", Expenses)
                cmd.Parameters.AddWithValue("@Edit", Edit)
                cmd.Parameters.AddWithValue("@ID", ID)


                'cmd.Parameters.AddWithValue("@ID", ExpensesID) ' Primary key for matching record  
                'cmd.ExecuteNonQuery()

                'Dim expense As New Class2
                'Class2.expense.Edit = TextBox1.Text
                'Dim txtfeedback As New TextBox()
                'txtfeedback.Name = "lblDateTime"
                'txtfeedback.Text = "Today: " & DateTime.Now.ToString("dddd, dd MMMM yyyy HH: ss")
                'txtfeedback.Location = New Point(20, 20) 'Adjust position as needed
                'txtfeedback.AutoSize = True
                'txtfeedback.Font = New Font("Arial", 20, FontStyle.Bold)
                'txtfeedback.ForeColor = Color.Blue
                'Me.Controls.Add(txtfeedback)

                MsgBox("Expense Updated Successfuly!", vbInformation, "Update Confirmation")
                PopulateComboBox()
                loadExpensesdataFromDatabase()
                'expense.ClearControls(Me)

            End Using
        Catch ex As OleDbException
            '    MessageBox.Show($"Error updating Expenses in database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Debug.WriteLine("Fail to Initialize")
            MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            If DataGridView1.SelectedRows.Count > 0 Then
                Debug.WriteLine("A Row is selected for upgrade.")
                'Rest of your edit Logic
            Else
                Debug.WriteLine("No Row selected, Exiting BtnEdit_Click.")
                MessageBox.Show("please select an expense to update.", "update Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            Debug.WriteLine("Existing BtnEdit successfully")
        End Try
    End Sub

    Private Sub BtnExport_Click(sender As Object, e As EventArgs) Handles BtnExport.Click


        Try
            Dim filePath As String = "C:\Users\user\Documents\Visual Studio 2010\Projects\My hello\HOUSEHOLD MANAGEMENT SYSTEM.vb1\HOUSEHOLD MANAGEMENT SYSTEM.vb1\Expenses.Txt.txt"
            Using writer As New StreamWriter(filePath, False)
                writer.WriteLine("Date,Amount,Currency,Tags,Frequency") ' Header

                For Each row As DataGridViewRow In DataGridView1.Rows
                    If Not row.IsNewRow Then
                        Dim dateValue = row.Cells("Dates").Value.ToString()
                        Dim amount = row.Cells("Amount").Value.ToString()
                        Dim currency = row.Cells("Currency").Value.ToString()
                        Dim tags = row.Cells("Tags").Value.ToString()
                        Dim frequency = row.Cells("Frequency").Value.ToString()

                        Dim line As String = $"{dateValue},{amount},{currency},{tags},{frequency}"
                        writer.WriteLine(line)
                    End If
                Next
            End Using
            MessageBox.Show($"Expenses exported successfully to {filePath}.", "Export Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            Debug.WriteLine($"Error exporting data: {ex.Message}")
            MessageBox.Show("An error occurred while exporting data.", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnImport_Click(sender As Object, e As EventArgs) Handles BtnImport.Click


        Try
            Dim filePath As String = "C:\Users\user\Documents\Visual Studio 2010\Projects\My hello\HOUSEHOLD MANAGEMENT SYSTEM.vb1\HOUSEHOLD MANAGEMENT SYSTEM.vb1\Expenses.ImportTxt.txt"


            If Not File.Exists(filePath) Then
                MessageBox.Show("File not found.", "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Using reader As New StreamReader(filePath)
                ' Skip the header
                reader.ReadLine()

                While Not reader.EndOfStream
                    Dim line As String = reader.ReadLine()
                    Dim values As String() = line.Split(","c)

                    ' Populate DataGridView
                    'DataGridView1.Rows.Add(exp.Amount, exp.Currency, exp.payment.ToString, exp.Tags, exp.frequency, exp.ApprovalStatus, exp, exp.Edit())
                End While
            End Using
            MessageBox.Show("Expenses imported successfully.", "Import Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            Debug.WriteLine($"Error importing data: {ex.Message}")
            MessageBox.Show("An error occurred while importing data.", "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnRefresh_Click(sender As Object, e As EventArgs) Handles BtnRefresh.Click


        Try
            Dim searchTerm As String = txtSearch.Text.ToLower()

            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not row.IsNewRow Then
                    Dim tags As String = row.Cells("Tags").Value.ToString().ToLower()
                    Dim frequency As String = row.Cells("Frequency").Value.ToString().ToLower()

                    row.Visible = tags.Contains(searchTerm) OrElse frequency.Contains(searchTerm)
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine($"Error in search functionality: {ex.Message}")
            'MessageBox.Show("An error occurred while searching.", "Search Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged

    End Sub

    Private Sub BtnSendEmail_Click(sender As Object, e As EventArgs)

    End Sub



    'Private Sub BtnSendEmail_Click(sender As Object, e As EventArgs) Handles BtnSendEmail.Click


    '    '' Define the low stock threshold
    '    'Dim lowStockThreshold As Integer = 100

    '    '' Get low stock items from database
    '    'Dim lowStockItems As List(Of String) = GetLowStockItems(lowStockThreshold)

    '    '' Send an email if any items are low on stock
    '    'If lowStockItems.Count > 0 Then
    '    '    Dim messageBody As String = "The following items are low in stock:" & vbCrLf & String.Join(vbCrLf, lowStockItems)
    '    '    SendEmail("dongolamaano3@gmail.com", "Expense Alert: Low Stock", messageBody)
    '    'Else
    '    '    MessageBox.Show("Expenses are sufficient.", "Expense Check", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '    'End If
    '    Dim expenseThreshold As Decimal = budgetLimit * 0.8D
    '    Dim expenseDetails As String = GetTotalExpensesWithTags()

    '    ' Extract the total expenses from the last line of the expenseDetails string.
    '    Dim totalExpenses As Decimal = 0
    '    Dim lines() As String = expenseDetails.Split({vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
    '    For Each line As String In lines
    '        If line.StartsWith("Total Expenses:") Then
    '            Dim expenseValueString As String = line.Replace("Total Expenses:", "").Trim()
    '            Decimal.TryParse(expenseValueString, Globalization.NumberStyles.Currency, Nothing, totalExpenses)
    '            Exit For
    '        End If
    '    Next

    '    If totalExpenses > expenseThreshold Then
    '        Dim messageBody As String = $"Alert! Expenses exceeded 80% of the budget.{vbCrLf}" &
    '        $"Budget Limit: {budgetLimit:C}{vbCrLf}{vbCrLf}" &
    '                                    $"Expense Breakdown:{vbCrLf}{expenseDetails}"
    '        SendEmail("austinmulalo113@gmail.com", "Budget Alert: High Expenses", messageBody)
    '    Else
    '        MessageBox.Show("Expenses are within budget.", "Budget Check", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '    End If
    'End Sub
    '' Declarations

    'Private connectionString As String = Module1.connectionString
    'Private budgetLimit As Decimal = 600
    'Private emailFrom As String = "austinmulalo113@gmail.com"
    'Private emailPassword As String = "oqsa qwqa bhjc nzoe" ' Use your Gmail/App Password here
    '' Retrieves expense details grouped by Tags and calculates the total
    'Private Function GetTotalExpensesWithTags() As String
    '    Dim totalExpenses As Decimal = 0
    '    Dim expenseDetails As New List(Of String)
    '    Try
    '        Using conn As New OleDbConnection(Module1.connectionString)

    '            conn.Open()
    '            Dim query As String = "SELECT Currency, Amount  FROM Expenses "
    '            Using cmd As New OleDbCommand(query, conn)
    '                Using reader As OleDbDataReader = cmd.ExecuteReader()
    '                    While reader.Read()
    '                        Dim tag As String = reader("Currency").ToString()
    '                        Dim amount As Decimal = Convert.ToDecimal(reader("Amount"))
    '                        totalExpenses += amount
    '                        expenseDetails.Add($"{tag}: {amount:C}")

    '                    End While
    '                End Using
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '        MessageBox.Show("Database error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    '    Dim expenseSummary As String = String.Join(vbCrLf, expenseDetails)
    '    Return $"{expenseSummary}{vbCrLf}Total Expenses: {totalExpenses:C}"
    'End Function
    'Private Sub SendEmail(toEmail As String, subject As String, body As String)

    '    Try
    '        Dim smtpClient As New SmtpClient("smtp.gmail.com") With {
    '                        .Port = 587,
    '                      .Credentials = New Net.NetworkCredential("austinmulalo113@gmail.com", "oqsa qwqa bhjc nzoe"), ' Replace with Gmail/App Password
    '                       .EnableSsl = True}

    '        Dim mail As New MailMessage() With {
    '            .From = New MailAddress("austinmulalo113@gmail.com"),
    '            .Subject = subject,
    '            .Body = body
    '        }

    '        mail.To.Add(toEmail)
    '        smtpClient.Send(mail)
    '        MessageBox.Show("Email sent successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '    Catch ex As Exception
    '        MessageBox.Show("Failed to send email: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub
End Class
