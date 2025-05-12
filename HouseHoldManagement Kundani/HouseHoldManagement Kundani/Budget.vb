Imports System.IO
Imports System.Data.OleDb
Public Class Budget
    Public Property conn As New OleDbConnection(connectionString)


    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Dim income, rent, utilities, groceries, otherExpenses As Decimal
        Dim totalExpenses, remaining As Decimal

        ' Validate and convert inputs
        If Not Decimal.TryParse(TextBox1.Text, income) Then
            MessageBox.Show("Enter a valid number for income.")
            Exit Sub
        End If
        Decimal.TryParse(TextBox2.Text, rent)
        Decimal.TryParse(TextBox3.Text, utilities)
        Decimal.TryParse(TextBox4.Text, groceries)
        'Decimal.TryParse(TextBox5.Text, otherExpenses)

        ' Calculate
        totalExpenses = rent + utilities + groceries + otherExpenses
        remaining = income - totalExpenses

        ' Display results
        Label6.Text = "Budget: R" & totalExpenses.ToString("F2")
        Label7.Text = "Remaining Balance: R" & remaining.ToString("F2")

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs)
        End
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub

    Public currentUsername As String
    Private Sub Budget_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        LoadBudgetDataFromDatabase()
        ToolTip1.SetToolTip(Button2, "Save")
        ToolTip1.SetToolTip(Button3, "Edit")
        ToolTip1.SetToolTip(Button4, "Delete")
        ToolTip1.SetToolTip(Button1, "Calculate")
        ToolTip1.SetToolTip(Button5, "Filter")
        ToolTip1.SetToolTip(Button6, "Sort")


    End Sub



    Private Sub LockForm()
        If TextBox7.Text = "Member" OrElse TextBox7.Text = "Chef" Then
            ' Lock the form

            MessageBox.Show("Access Denied: Only Finance role can access the Budget form.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Me.Enabled = False
        ElseIf TextBox7.Text = "Finance" OrElse TextBox7.Text = "Admin" Then
            ' Unlock the form
            Me.Enabled = True
        Else
            ' Optional: handle other cases
            Me.Enabled = False
        End If
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        LockForm()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Debug.WriteLine("Entering btnSubmit")
        Try

            Debug.WriteLine("User confirmed btnSubmit")
            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                conn.Open()

                ' Update the table name if necessary  
                Dim tableName As String = "Budget"

                ' Create an OleDbCommand to insert the Budget data into the database 
                Dim cmd As New OleDbCommand("INSERT INTO [Budget] ([Person], [Role], [Frequency], [Income], [Utilities], [Groceries], [Expenses], [StartDate], [EndDate]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)", conn)

                ' Set the parameter values from the UI controls 
                'Class declaretions

                Dim budgets As New BudgetClass With {
                    .Person = TextBox6.Text,
                    .Role = TextBox7.Text,
                    .Frequency = ComboBox1.SelectedItem.ToString,
                    .Income = TextBox1.Text,
                    .Expenses = TextBox2.Text,
                    .Utilities = TextBox3.Text,
                    .Groceries = TextBox4.Text,
                    .StartDate = DateTimePicker1.Value,
                          .EndDate = DateTimePicker2.Value}

                cmd.Parameters.Clear()


                'cmd.Parameters.AddWithValue("@ExpenseID", expense.ExpenseID)
                cmd.Parameters.AddWithValue("@Person", budgets.Person)
                cmd.Parameters.AddWithValue("@Role", budgets.Role)
                cmd.Parameters.AddWithValue("@Frequency", budgets.Frequency)
                cmd.Parameters.AddWithValue("@Income", budgets.Income)
                ' cmd.Parameters.AddWithValue("@Rent", budgets.Rent)
                cmd.Parameters.AddWithValue("@Utilities", budgets.Utilities)
                cmd.Parameters.AddWithValue("@Groceries", budgets.Groceries)
                cmd.Parameters.AddWithValue("@Expenses", budgets.Expenses)
                cmd.Parameters.AddWithValue("@StartDate", budgets.StartDate)

                cmd.Parameters.AddWithValue("@EndDate", budgets.EndDate)


                MsgBox("Expense Information Saved!" & vbCrLf &
                        "Person: " & budgets.Person & vbCrLf &
                        "Role: " & budgets.Role & vbCrLf &
                        "Frequency: " & budgets.Frequency & vbCrLf &
                        "Income: " & budgets.Income & vbCrLf &
                        "Utilities: " & budgets.Utilities & vbCrLf &
                        "Groceries: " & budgets.Groceries & vbCrLf &
                         "Expenses: " & budgets.Expenses & vbCrLf &
                         "StartDate: " & budgets.StartDate & vbCrLf &
                         "EndDate: " & budgets.EndDate.ToString, vbInformation, "Budget Confirmation")

                cmd.ExecuteNonQuery()

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
        LoadBudgetDataFromDatabase()
        Debug.WriteLine("Exiting btnSubmit")

    End Sub
    Public Sub LoadBudgetDataFromDatabase()
        Try
            '  Dim dataTable As DataTable = HouseHold.GetData("SELECT * FROM Budget")
            ' DataGridView1.DataSource = DataTable
            Debug.WriteLine("Populate Datagridview: Datagridview populated successfully.")
            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                conn.Open()

                Dim tableName As String = "Budget"

                Dim cmd As New OleDbCommand($"SELECT*FROM {tableName}", conn)

                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)
                DataGridView1.DataSource = dt
            End Using
        Catch ex As Exception
            Debug.WriteLine("Failed to populate Gatagridview")
            Debug.Write($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error Loading Budget data from database: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Dim income, rent, utilities, groceries, otherExpenses As Decimal
        Dim totalExpenses, remaining As Decimal

        ' Validate and convert inputs
        If Not Decimal.TryParse(TextBox1.Text, income) Then
            MessageBox.Show("Enter a valid number for income.")
            Exit Sub
        End If
        Decimal.TryParse(TextBox2.Text, otherExpenses)
        Decimal.TryParse(TextBox3.Text, utilities)
        Decimal.TryParse(TextBox4.Text, groceries)
        ' Decimal.TryParse(TextBox5.Text, otherExpenses)

        ' Calculate
        totalExpenses = rent + utilities + groceries + otherExpenses
        remaining = income - totalExpenses

        ' Display results
        Label6.Text = "Total Expenses: R" & totalExpenses.ToString("F2")
        Label7.Text = "Remaining Balance: R" & remaining.ToString("F2")

        If remaining < 0 Then
            Label7.ForeColor = Color.Black
            Timer1.Start()
        Else
            Label7.ForeColor = Color.Black
            Timer1.Stop()
            Label7.Visible = True ' Ensure it's visible when not blinking
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Debug.WriteLine("Entering button update click")
        If DataGridView1.SelectedRows.Count = 0 Then
            Debug.WriteLine("User confirmed update")
            MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try

            Debug.WriteLine($"Format error in button update:")
            Debug.WriteLine("Updating data: data updated")
            Dim Person As String = TextBox6.Text
            Dim Role As String = TextBox7.Text
            Dim Frequency As String = ComboBox1.SelectedItem.ToString
            Dim Income As String = TextBox1.Text
            'Dim Rent As String = TextBox2.Text
            Dim Utilities As String = TextBox3.Text
            Dim Groceries As String = TextBox4.Text
            Dim Expenses As String = TextBox2.Text
            Dim StartDate As String = DateTimePicker1.Value
            Dim EndDate As String = DateTimePicker2.Value


            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)

                conn.Open()

                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Create an OleDbCommand to update the personnel data in the database  
                Dim cmd As New OleDbCommand("UPDATE [Budget] ([Person], [Role], [Frequency], [Income], [Utilities], [Groceries], [Expenses], [StartDate], [EndDate]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", conn)

                'cmd.Parameters.AddWithValue("@ExpenseID", expense.ExpenseID)
                cmd.Parameters.AddWithValue("@Person", Person)
                cmd.Parameters.AddWithValue("@Role", Role)
                cmd.Parameters.AddWithValue("@Frequency", Frequency)
                cmd.Parameters.AddWithValue("@Income", Income)
                cmd.Parameters.AddWithValue("@Expenses", Expenses)
                cmd.Parameters.AddWithValue("@Utilities", Utilities)
                cmd.Parameters.AddWithValue("@Groceries", Groceries)
                'cmd.Parameters.AddWithValue("@Expenses", Expenses)
                cmd.Parameters.AddWithValue("@StartDate", StartDate)
                cmd.Parameters.AddWithValue("@EndDate", EndDate)


                MsgBox("Inventory Items Updated Successfuly!", vbInformation, "Update Confirmation")
                LoadBudgetDataFromDatabase()
                '  InventoryModule.ClearControls(Me)

            End Using

        Catch ex As OleDbException
            Debug.WriteLine("User cancelled update")
            Debug.WriteLine("Unexpected error in button update")
            Debug.Write($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error updating inventory in database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Debug.WriteLine("Exiting button update")

    End Sub

    'Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
    '    Dim userInput As String = TextBox7.Text.Trim().ToLower()

    '    If userInput ISNOT "Admin" AndAlso userInput ISNOT "Finance" Then
    '        ' Disable all controls except TextBox1
    '        For Each ctrl As Control In Me.Controls
    '            If ctrl IsNot TextBox2 Then
    '                ctrl.Enabled = False
    '            End If
    '        Next

    '        MessageBox.Show("Access restricted. Only 'Admin' or 'Finance' can proceed.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '    End If
    'End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Debug.WriteLine("Entering button delete")
        ' Check if there are any selected rows in the DataGridView for expenses  
        If DataGridView1.SelectedRows.Count > 0 Then

            ' Get the selected row  
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Assuming there is an "ID" column which is the primary key in the database  
            Dim PhotoID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  
            'Dim DeletedBy As String

            ' Confirm deletion  
            Dim confirmationResult As DialogResult = MessageBox.Show("Are you sure you want to delete this Budget?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If confirmationResult = DialogResult.Yes Then
                Debug.WriteLine("User confirmation deletion.")
                ' Proceed with deletion  
                Try
                    Debug.WriteLine("Format errors in button delete")
                    Debug.WriteLine("Deleting data: Data delected")
                    Debug.WriteLine("Stack Trace: {ex.StackTrace}")
                    Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                        conn.Open()

                        ' Create the delete command  
                        Dim cmd As New OleDbCommand("DELETE FROM [Budget] WHERE [ID] = ?", conn)
                        cmd.Parameters.AddWithValue("@ID", PhotoID) ' Primary key for matching record  

                        ' Execute the delete command  
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show("Budget deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ' Optionally refresh DataGridView or reload from database  
                            ' PopulateDataGridView()
                        Else
                            MessageBox.Show("No Photo was deleted. Please check if the ID exists.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End Using
                    LoadBudgetDataFromDatabase()
                Catch ex As Exception
                    Debug.WriteLine("Failed to delete data")
                    Debug.Write($"Stack Trace: {ex.StackTrace}")
                    MessageBox.Show($"An error occurred while deleting the Budget: {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            Debug.WriteLine("User cancelled deletion")
            MessageBox.Show("Please select Budget to delete.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

        Debug.WriteLine("Exiting button delete")

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim selectedFrequency As String = If(ComboBox1.SelectedItem IsNot Nothing, ComboBox1.SelectedItem.ToString(), "")
        '  Dim selectedUnit As String = If(ComboBox2.SelectedItem IsNot Nothing, ComboBox2.SelectedItem.ToString(), "")

        HouseHoldManagment_Module.FilterBudget(selectedFrequency) ', selectedUnit)
    End Sub

    Private Sub Button6_Click_1(sender As Object, e As EventArgs) Handles Button6.Click
        DataGridView1.Sort(DataGridView1.Columns("StartDate"), System.ComponentModel.ListSortDirection.Ascending)
    End Sub

    'BLINK
    ' Declare a variable to track the label visibility
    Private labelBlinkState As Boolean = True
    Private Sub CheckBalanceAndBlink()
        ' Assume remainingBalance is a variable storing your current balance
        Dim remainingBalance As Decimal = Decimal.Parse(Label7.Text)

        If remainingBalance < 0 Then
            Timer1.Start()
        Else
            Timer1.Stop()
            Label7.Visible = True ' Reset visibility
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label7.Visible = labelBlinkState
        labelBlinkState = Not labelBlinkState
    End Sub
End Class