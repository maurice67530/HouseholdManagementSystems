
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Net.Mail
Imports System.Net



Public Class MealPlan
    Private row As Object
    Private currentRowIndex As Integer = 0
    Private mealPlanData As DataTable

    Private Sub MealPlan_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        LoadMealPlanFromDatabase()
        'PopulateListboxFromdatabase()

        'Clear any existing items in ComboBox1
        ComboBox3.Items.Clear()

        Using conn As New OleDbConnection(Module1.connectionString)
            conn.Open()

            ' Query to fetch all ItemName values from Inventory1
            Dim fetchcommand As New OleDbCommand("SELECT ItemName FROM Inventory", conn)


            Using Readers As OleDbDataReader = fetchcommand.ExecuteReader()
                ' Add ItemName values to ComboBox1
                While Readers.Read()
                    ComboBox3.Items.Add(Readers("ItemName").ToString())
                End While


                ' Set tooltips For buttons
                ToolTip1.SetToolTip(Button1, "Dashboard")
                ToolTip1.SetToolTip(Button2, "Save")
                ToolTip1.SetToolTip(Button3, "Edit")
                ToolTip1.SetToolTip(Button4, "Delete")
                ToolTip1.SetToolTip(Button5, "Sort")
                ToolTip1.SetToolTip(Button6, "Filter")
                ToolTip1.SetToolTip(Button7, "Highlight")
                ToolTip1.SetToolTip(Button8, "Calculate")




                ComboBox4.Items.Add("Day")
                ComboBox4.Items.Add("Week")
                ComboBox4.Items.Add("Month")
                ComboBox4.SelectedItem = 0

                ' Define the low stock threshold
                Dim lowStockThreshold As Integer = 3

                ' Get low stock items from database
                Dim lowStockItems As List(Of String) = GetLowStockItems(lowStockThreshold)

                ' Send an email if any items are low on stock
                If lowStockItems.Count > 0 Then
                    Dim messageBody As String = "The following items are low in stock:" & vbCrLf & String.Join(vbCrLf, lowStockItems)
                    SendEmail("austinmulalo113@gmail.com", "Inventory Alert: Low Stock", messageBody)
                Else
                    MessageBox.Show("Stock levels are sufficient.", "Inventory Check", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        End Using

    End Sub

    Private Function GetLowStockItems(threshold As Integer) As List(Of String)
        Dim lowStockItems As New List(Of String)

        Try
            Using conn As New OleDbConnection(connectionString)
                conn.Open()
                Dim query As String = "SELECT ItemName, Quantity FROM Inventory WHERE Quantity < @Threshold"
                Using cmd As New OleDbCommand(query, conn)
                    cmd.Parameters.AddWithValue("@Threshold", threshold)

                    Using reader As OleDbDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            lowStockItems.Add($"{reader("ItemName")}: {reader("Quantity")} remaining")
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Database error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return lowStockItems
    End Function

    Private Sub SendEmail(toEmail As String, subject As String, body As String)
        Try
            Dim smtpClient As New SmtpClient("smtp.gmail.com") With {
                            .Port = 587,
                            .Credentials = New Net.NetworkCredential("austinmulalo113@gmail.com", "oqsa qwqa bhjc nzoe"), ' Replace with Gmail/App Password
                            .EnableSsl = True
                        }

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


    'Private currentRowIndex As Integer = 0

    Private Sub LoadFilteredMealPlan()
        Using dbConnection As New OleDbConnection(connectionString)
            Dim selectedFilter As String = ComboBox4.SelectedItem?.ToString()
            Dim query As String = "SELECT * FROM MealPlan WHERE filter = ? AND 1=1"
            Dim startDate As Date = Date.Today
            Dim endDate As Date = Date.Today

            If Not String.IsNullOrEmpty(selectedFilter) Then
                Select Case selectedFilter
                    Case "Day"
                        endDate = startDate
                    Case "Week"
                        endDate = startDate.AddDays(7)
                    Case "Month"
                        endDate = startDate.AddMonths(1)
                End Select
            End If

            ' Create command with parameters
            Dim cmd As New OleDbCommand(query, dbConnection)
            cmd.Parameters.AddWithValue("?", selectedFilter) ' Bind Frequency value

            Dim adapter As New OleDbDataAdapter(cmd)
            mealPlanData = New DataTable()
            adapter.Fill(mealPlanData)

            DataGridView1.DataSource = mealPlanData ' Display filtered data in DataGridView
        End Using
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs)
        regis.ShowDialog()

    End Sub

    Private Sub Button1_Click_1(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        DashBoard.Show()
        Me.Close()
    End Sub
    'Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
    '    If DataGridView1.Columns(e.ColumnIndex).Name = "EndDate" Then
    '        Dim expiryDate As Date = CDate(e.Value)
    '        If expiryDate < Date.Now Then
    '            e.CellStyle.BackColor = Color.Red
    '            e.CellStyle.ForeColor = Color.White
    '        End If
    '    End If
    'End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        'If lstSuggestedMeals.SelectedItem Is Nothing Then
        '    MsgBox("Please select a meal to save.", MsgBoxStyle.Exclamation, "Select Meal")
        '    Exit Sub
        'End If

        'Dim selectedMeal As String = lstSuggestedMeals.SelectedItem.ToString()

        'Try
        Debug.WriteLine("Entering btnSubmit")
        Dim Meal As New MealPlan With {
        .mealPlanID = TextBox1.Text,
       .NameOfMeal = TextBox2.Text,
       .StartDate = DateTimePicker1.Text,
       .MealsType = ComboBox1.SelectedItem,
      .EndDate = DateTimePicker2.Text,
      .totalcalories = ComboBox2.Text,
       .Description = TextBox3.Text,
       .Picture = PictureBox4.Text}
        '.filter = ComboBox4.SelectedItem}

        Using connect As New OleDbConnection(Module1.connectionString)
            conn.Open()

            Dim tablename As String = "MealPlan"

            Dim cmd As New OleDbCommand("INSERT INTO [MealPlan] ([MealPlanID],[NameOfMeal],[StartDate],[MealsType],[EndDate],[totalcalories],[Description],[Picture]) VALUES (@MealPlanID ,@NameOfMeal ,@StartDate, @MealsType, @EndDate, @Totalcalories, @Description, @Picture)", Module1.conn)

            'params
            cmd.Parameters.AddWithValue("MealPlanID", Meal.mealPlanID)
            cmd.Parameters.AddWithValue("NameOfMeal", Meal.NameOfMeal)
            cmd.Parameters.AddWithValue("StartDate", Meal.StartDate)
            cmd.Parameters.AddWithValue("MealsType", Meal.MealsType)
            cmd.Parameters.AddWithValue("EndDate", Meal.EndDate)
            cmd.Parameters.AddWithValue("totalcalories", Meal.totalcalories)
            cmd.Parameters.AddWithValue("Description", Meal.Description)
            cmd.Parameters.AddWithValue("picture", Meal.Picture)
            'cmd.Parameters.AddWithValue("filter", Meals.filter)
            cmd.ExecuteNonQuery()


            MsgBox("MealPlan  Added!" & vbCrLf &
             "MealPlanID: " & Meal.mealPlanID & vbCrLf &
               "NameOfMeal: " & Meal.NameOfMeal.ToString & vbCrLf &
               "StartDate: " & Meal.StartDate & vbCrLf &
               "MealsType: " & Meal.MealsType & vbCrLf &
               "EndDate: " & Meal.EndDate & vbCrLf &
               "totalcalories: " & Meal.totalcalories & vbCrLf &
               "Description: " & Meal.Description & vbCrLf &
               "Picture:" & Picture, vbInformation, "MealPlan Confirmation")
            '"Picture:" & Picture, vbInformation, "MealPlan Confirmation")
            ' '"filter: " & filter, vbInformation, "MealPlan Confirmation")

            conn.Close()
        End Using


        Dim selectedItem As String = ComboBox1.SelectedItem() ' Meal selected from ComboBox2
        Using conn As New OleDbConnection(Module1.connectionString)

            conn.Open()

            '' Fetch the item and its details from the Inventory1 table.
            'Dim fetchcommand As New OleDbCommand("SELECT ItemName, Quantity, PricePerItem FROM Inventory WHERE ItemName = ?", conn)
            'fetchcommand.Parameters.AddWithValue("@ItemName", selectedItem)

            'Using Readers As OleDbDataReader = fetchcommand.ExecuteReader()
            '    If Readers.Read() Then

            '    End If
            '    Dim ItemQuantity As Integer = Convert.ToInt32(Readers("Quantity")) ' Get the available total quantity of the item

            '    ' If the item is in stock
            '    If ItemQuantity > 0 Then

            '    End If
            '    ' Update the inventory by reducing quantity by 1 (the meal uses one unit of the item)
            '    Dim updateCommand As New OleDbCommand("UPDATE Inventory SET Quantity = Unit - 1 WHERE ItemName = ?", conn)
            '    updateCommand.Parameters.AddWithValue("@ItemName", selectedItem)
            '    updateCommand.ExecuteNonQuery()

            '    ' Display confirmation message
            '    MessageBox.Show("Item added to Meal Plan. Inventory updated.")

            '    ' Display updated stock status
            '    TextBox4.Text = "Available Stock: " & (ItemQuantity - 1).ToString()


            '    ' Item is out of stock
            '    MessageBox.Show("Item is out of stock.")
            '    TextBox4.Text = "Available Stock: 0"
            '     TextBox4.Text = ""

            ' Fetch the item and its details from the Inventory1 table.
            Dim fetchcommand As New OleDbCommand("SELECT ItemName, Quantity FROM Inventory WHERE ItemName = ?", conn)
            fetchcommand.Parameters.AddWithValue("@ItemName", selectedItem)

            Using Readers As OleDbDataReader = fetchcommand.ExecuteReader()
                If Readers.Read() Then

                End If
                Dim ItemQuantity As Integer = Convert.ToInt32(Readers("Quantity")) ' Get the available total quantity of the item

                ' If the item is in stock
                If ItemQuantity > 0 Then


                    ' Update the inventory by reducing quantity by 1 (the meal uses one unit of the item)
                    Dim updateCommand As New OleDbCommand("UPDATE Inventory SET Quantity = Quantity - 1 WHERE ItemName = ?", conn)
                    updateCommand.Parameters.AddWithValue("@ItemName", selectedItem)
                    updateCommand.ExecuteNonQuery()



                    ' Display updated stock status
                    TextBox4.Text = "Available Stock: " & (ItemQuantity - 1).ToString()

                Else

                    ' Item is out of stock
                    MessageBox.Show("Item is out of stock.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    TextBox4.Text = "Available Stock: 0"
                    TextBox4.Text = " "

                End If



            End Using
        End Using
        Try

        Catch ex As OleDbException
            Debug.WriteLine($"Database error: {ex.Message}")
            Debug.WriteLine($"stack Trace:{ex.StackTrace}")
            MessageBox.Show("Error saving test to database. Please check the connection.", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Debug.WriteLine($"General error: {ex.Message}")
            MessageBox.Show("An unexpected error occurred.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            LoadMealPlanFromDatabase()
            conn.Close()

        End Try

    End Sub
    Public Sub LoadMealPlanFromDatabase()
        Try
            Debug.WriteLine("populated connected succesfully")
            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()

                Dim tableName As String = "MealPlan"
                Dim cmd As New OleDbCommand($"SELECT * FROM {tableName}", conn)

                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)

                DataGridView1.DataSource = dt
                'HighlightExpiredItemss()
            End Using

        Catch ex As Exception

            Debug.WriteLine("populated fail to initialize succesfully")
            MessageBox.Show($"Error Loading  Expenses data from database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' Check if there are any selected rows in the DataGridView for expenses  
        If DataGridView1.SelectedRows.Count > 0 Then
            ' Get the selected row  
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Assuming there is an "ID" column which is the primary key in the database  
            Dim TaskManagementId As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  

            ' Confirm deletion  
            Dim confirmationResult As DialogResult = MessageBox.Show("Are you sure you want to delete this expense?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If confirmationResult = DialogResult.Yes Then
                ' Proceed with deletion  
                Try
                    Using conn As New OleDbConnection(Module1.connectionString)
                        conn.Open()

                        ' Create the delete command  
                        Dim cmd As New OleDbCommand("DELETE FROM [MealPlan] WHERE [ID] = ?", conn)
                        cmd.Parameters.AddWithValue("@ID", TaskManagementId) ' Primary key for matching record  

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

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        'Load the data from  the selected row into UI controls 
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
            TextBox1.Text = selectedRow.Cells("MealPlanID").Value.ToString()
            TextBox2.Text = selectedRow.Cells("NameOfMeal").Value.ToString()
            DateTimePicker1.Text = selectedRow.Cells("StartDate").Value.ToString()
            ComboBox1.Text = selectedRow.Cells("MealsType").Value.ToString()
            DateTimePicker2.Text = selectedRow.Cells("EndDate").Value.ToString()
            ComboBox2.Text = selectedRow.Cells("totalcalories").Value.ToString()
            TextBox3.Text = selectedRow.Cells("Description").Value.ToString()
            PictureBox4.Text = selectedRow.Cells("Picture").Value.ToString()
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Debug.WriteLine("Entering btnEdit")
            Dim MealPlanID As Integer = TextBox1.Text
            Dim NameOfMeal As String = TextBox2.Text
            Dim StartDate As DateTime = DateTimePicker1.Text
            Dim MealsType As String = TextBox3.Text
            Dim EndDate As DateTime = DateTimePicker2.Text
            Dim TotalCalories As String = ComboBox2.Text
            Dim Description As String = TextBox3.Text
            Dim Picture As String = PictureBox4.Text
            Dim filter As String = ComboBox4.Text



            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()

                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Create an OleDbCommand to update the expense data in the database  

                Dim cmd As New OleDbCommand("UPDATE MealPlan SET [MealPlanID] = ?, [NameOfMeal]  = ?, [StartDate] = ?, [Meals] = ?, [EndDate] = ?, [totalCalories] = ?,[Description]=?,[Picture]=?, [filter] =? WHERE [ID] = ?", conn)
                'Set the parameter values from the UI controls 
                cmd.Parameters.AddWithValue("@MealPlanID", TextBox1.Text)
                cmd.Parameters.AddWithValue("@NameOfMeal", TextBox2.Text)
                cmd.Parameters.AddWithValue("@StartDate", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@Meals", TextBox3.Text)
                cmd.Parameters.AddWithValue("@EndDate", DateTimePicker2.Text)
                cmd.Parameters.AddWithValue("@TotalCalories", ComboBox2.Text)
                cmd.Parameters.AddWithValue("@Description", TextBox3.Text)
                cmd.Parameters.AddWithValue("@Picture", ComboBox1.Text)
                cmd.Parameters.AddWithValue("@filter", ComboBox4.Text)

                cmd.Parameters.AddWithValue("@ID", MealsType) ' Primary key for matching record  
                cmd.ExecuteNonQuery()
                MsgBox("MealPlan Updated Successfuly!", vbInformation, "Update Confirmation")

            End Using
        Catch ex As OleDbException
            MessageBox.Show($"Error updating MealPlan in database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        conn.Close()
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        '    Dim Tasktable As DataTable()

        '    Try
        '        Dim query As String = "SELECT *FROM CHORES WHERE 1=1"
        '        If Not String.IsNullOrEmpty(Meals) Then
        '            query &= "And Frequency =@Meals"
        '        End If

        '        If Not String.IsNullOrEmpty(totalcalories) Then
        '            query &= "And Priority =@totalcalories"
        '        End If
        '        Dim command As New OleDbCommand(query, conn)
        '        If Not String.IsNullOrEmpty(Meals) Then
        '            command.Parameters.AddWithValue("@meals", Meals)
        '        End If
        '        If Not String.IsNullOrEmpty(totalcalories) Then
        '            command.Parameters.AddWithValue("@totalcalories", totalcalories)

        '        End If

        '        Dim Adapter As New OleDbDataAdapter(command)
        '        Adapter.Fill(Tasktable)
        '        Chores.DataGridView1.DataSource = Tasktable

        '    Catch ex As Exception

        '        MsgBox("Error Filtering tasks " & ex.Message, MsgBoxStyle.Critical, "DataBase error")
        '        Debug.WriteLine($"stack Trace:{ex.StackTrace}")
        '    Finally
        '        conn.Close()
        '    End Try

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        'Try

        '    For Each Rows As DataGridView In DataGridView1.Rows
        '        If row.Cells("DueDate").Value IsNot Nothing Then
        '            Dim DueDate As DateTime = Convert.ToDateTime(row.Cells("DueDate").Value)
        '            Dim Status As String = row.Cells("Status").Value.ToString

        '            If DueDate < DateTime.Now AndAlso Status <> "Completed" Then
        '                Rows.DefaultCellStyle.BackColor = Color.Red
        '            Else

        '                Rows.DefaultCellStyle.BackColor = Color.WhiteSmoke

        '            End If

        '        End If
        '    Next

        '    Dim IncompletedCount As Integer = 0
        '    For Each row As DataGridView In DataGridView1.Rows
        '        If row.SelectedCells("Status").Value IsNot Nothing AndAlso row.SelectedCells("Status").Value.ToString() <> "Completed" Then

        '            IncompletedCount += 1



        '        End If
        '    Next

        '    Label6.Text = "IncompletedCount chores:" & IncompletedCount.ToString
        'Catch ex As Exception
        '    MessageBox.Show("Error highlighting overdue chores")
        '    Debug.WriteLine($"stack Trace:{ex.StackTrace}")
        'End Try

        'For Each row As DataGridViewRow In DataGridView1.Rows
        '    Dim Startdate As Date = Convert.ToDateTime(row.Cells("StartDate").Value
        '    Dim EndDate As Date As 

        'Next

    End Sub

    Private Sub listBox1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub
    'Private Sub PopulateListboxFromdatabase()


    '    Dim conn As New OleDbConnection(Module1.connectionString)
    '    Try
    '        Debug.WriteLine("Listbox populated from database")
    '        conn.Open()

    '        'Retrieve the status and title columns from the chore tabel
    '        Dim query As String = "SELECT Meals,TotalCalories From MealPlan "
    '        Dim cmd As New OleDbCommand(query, conn)
    '        Dim reader As OleDbDataReader = cmd.ExecuteReader

    '        'bind the retrieved data to the combobox 
    '        listBox1.Items.Clear()
    '        While reader.Read()
    '            ListBox2.Items.Add($"{reader("Meals")}{reader("Totalcalories")}")
    '        End While

    '        'Close the database
    '        reader.Close()

    '    Catch ex As Exception
    '        Debug.WriteLine("Failed to populate listbox")
    '        Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
    '        'handle any exeptions that may occur
    '        MessageBox.Show($"Error:{ex.Message}")

    '    Finally

    '        'close the database connection 
    '        If conn.State = ConnectionState.Open Then
    '            conn.Close()
    '        End If
    '    End Try
    '    Debug.WriteLine("Done with populating listbox from database ")
    'End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        Try
            Debug.WriteLine("Entering btnCalculate")
            'Calculate total expenses  
            Dim totalcalories As Decimal = Decimal.Parse(ComboBox2.Text)

            ' Calculate average expenses based on frequency  
            Dim averageExpenses As Decimal = 0
            If ComboBox2.SelectedItem IsNot Nothing Then
                Select Case ComboBox2.Text()
                    Case "Monthly"
                        averageExpenses = totalcalories / 12
                    Case "Weekly"
                        averageExpenses = totalcalories / 52
                    Case "Daily"
                        averageExpenses = totalcalories / 365
                    Case Else
                        ' Handle other frequencies as needed  
                End Select
            End If


            ' Display total and average expenses on the form  
            Label12.Text = $"Totalcalories: R {totalcalories:N2}"
            Label6.Text = $"Averagecalories per {ComboBox2.Text}R{averageExpenses:N2}"


            'Existing calculation code...
            Debug.WriteLine($"calculation complete Total:{totalcalories}, Avarage:{averageExpenses}")
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

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged


        'Dim selectedItem As String = ComboBox3.SelectedItem.ToString() ' Meal selected from ComboBox2

        'Using conn As New OleDbConnection(Module1.connectionString)
        '    conn.Open()

        '    ' Fetch the item and its details from Inventory1 table.
        '    Dim fetchcommand As New OleDbCommand("SELECT ItemID, Unit, PricePerItem FROM Inventory WHERE ItemID = ?", conn)
        '    fetchcommand.Parameters.AddWithValue("@ItemID", selectedItem)

        '    Using Readers As OleDbDataReader = fetchcommand.ExecuteReader()
        '        If Readers.Read() Then
        '            Dim ItemQuantity As Integer = Convert.ToInt32(Readers("Unit")) ' Get available total quantity of the item
        '            Dim ItemTotalCost As Integer = Convert.ToInt32(Readers("PricePerItem")) ' Get total cost for the item

        '            ' If the item is in stock
        '            If ItemQuantity > 0 Then
        '                ' Display the calories or total cost for the selected item
        '                TextBox4.Text = "PricePerItem: " & ItemTotalCost.ToString()


        '                ' If stock is below 6, show a warning
        '                If ItemQuantity < 6 Then
        '                    MessageBox.Show("Warning: Stock is below 6 for this item.")
        '                End If

        '                ' Display the current available stock
        '                TextBox4.Text = "Available Stock: " & ItemQuantity.ToString()
        '            Else
        '                ' Item is out of stock
        '                MessageBox.Show("Item is out of stock.")
        '                TextBox4.Text = "Available Stock: 0"
        '            End If
        '        End If
        '    End Using
        'End Using

        'Dim selectedItem As String = ComboBox1.SelectedItem.ToString() ' Meal selected from ComboBox2

        'Using conn As New OleDbConnection(connectionString)
        '    conn.Open()

        '    ' Fetch the item and its details from Inventory1 table.
        '    Dim fetchcommand As New OleDbCommand("SELECT ItemName, Quantity, Unit FROM GroceryItemss WHERE ItemName = ?", conn)
        '    fetchcommand.Parameters.AddWithValue("@ItemName", selectedItem)

        '    Using Readers As OleDbDataReader = fetchcommand.ExecuteReader()
        '        If Readers.Read() Then
        '            Dim ItemQuantity As Integer = Convert.ToInt32(Readers("Quantity")) ' Get available total quantity of the item
        '            Dim ItemTotalCost As Integer = Convert.ToInt32(Readers("Unit")) ' Get total cost for the item

        '            ' If the item is in stock
        '            If ItemQuantity > 0 Then
        '                ' Display the calories or total cost for the selected item
        '                TextBox6.Text = "Unit: " & ItemTotalCost.ToString()
        '                Button1.Enabled = True

        '                'Dim selectedMeal As String = ComboBox1.SelectedItem.ToString()
        '                'GetMealSuggestion(selectedMeal)

        '                ' If stock is below 6, show a warning
        '                If ItemQuantity < 5 Then
        '                    MsgBox("Warning: Stock is below 5 for this item.", MessageBoxButtons.OK + MessageBoxIcon.Warning)
        '                    Button1.Enabled = True

        '                End If

        '                ' Display the current available stock
        '                TextBox6.Text = "Available Stock: " & ItemQuantity.ToString()

        '            Else
        '                ' Item is out of stock
        '                MsgBox("Item is out of stock.", MessageBoxButtons.OK + MessageBoxIcon.Error)
        '                TextBox6.Text = "Available Stock: 0"

        '                Button1.Enabled = False
        '            End If
        '        End If

        '    End Using
        'End Using
        Dim selectedItem As String = ComboBox3.SelectedItem.ToString() ' Meal selected from ComboBox2

        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            ' Fetch the item and its details from Inventory1 table.
            Dim fetchcommand As New OleDbCommand("SELECT ItemName,Quantity,Unit,ExpiryDate FROM GroceryItem WHERE ItemName = ?", conn)
            fetchcommand.Parameters.AddWithValue("@ItemName", selectedItem)

            Using Readers As OleDbDataReader = fetchcommand.ExecuteReader()
                If Readers.Read() Then
                    Dim ItemQuantity As Integer = Convert.ToInt32(Readers("Quantity")) ' Get available total quantity of the item
                    Dim ItemTotalCost As Integer = Convert.ToInt32(Readers("Unit")) ' Get total cost for the item
                    Dim ExpirationDate As DateTime = Convert.ToDateTime(Readers("ExpiryDate")) ' Get the expiration date of the item

                    ' Check if the item is expired
                    If ExpirationDate < DateTime.Now Then
                        MsgBox("This ingredient has expired. Unable to suggest a meal. You can select a fresh available Grocery Item to prepare a meal ", MessageBoxButtons.OK + MessageBoxIcon.Warning)
                        Button1.Enabled = False
                        TextBox4.BackColor = Color.Red ' Change text color to red
                        TextBox4.Text = "Item has expired."
                        Return ' Exit the method
                    End If

                    ' If the item is in stock
                    If ItemQuantity > 0 Then
                        ' Display the cost for the selected item
                        TextBox4.Text = "Unit: " & ItemTotalCost.ToString()
                        Button1.Enabled = True
                        TextBox4.BackColor = Color.PaleVioletRed ' Change text color to red
                        ' If stock is below 6, show a warning
                        If ItemQuantity < 5 Then
                            MsgBox("Warning: Stock is below 5 for this item.", MessageBoxButtons.OK + MessageBoxIcon.Warning)
                            Button1.Enabled = True
                            TextBox4.BackColor = Color.PaleVioletRed ' Change text color to red
                            TextBox4.Text = "Item has expired."
                        End If

                        ' Display the current available stock
                        TextBox4.Text = "Available Stock: " & ItemQuantity.ToString()

                    Else
                        ' Item is out of stock
                        MsgBox("Item is out of stock.", MessageBoxButtons.OK + MessageBoxIcon.Error)
                        TextBox4.Text = "Available Stock: 0"
                        Button1.Enabled = True
                        TextBox4.BackColor = Color.Red
                    End If
                Else
                    ' Handle the case where the item is not found
                    MsgBox("Item not found.", MessageBoxButtons.OK + MessageBoxIcon.Error)
                End If
            End Using
        End Using

        'Module1.SendEmail()
    End Sub
    Private Sub BtnPrintMealPlan_Click(sender As Object, e As EventArgs) Handles BtnPrintMealPlan.Click

        PrintDialog1.Document = PrintDocument1
        If PrintDialog1.ShowDialog() = DialogResult.OK Then
            LoadFilteredMealPlan() ' Load filtered data based on selected frequency
            If mealPlanData.Rows.Count > 0 Then
                PrintDocument1.Print()
            Else
                MessageBox.Show("No meal plans found for the selected period.", "Print Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End If
    End Sub
    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim font As New Font("Arial", 12)
        Dim brush As New SolidBrush(Color.Black)
        Dim yPos As Integer = 100
        Dim leftMargin As Integer = e.MarginBounds.Left

        ' Print Title
        e.Graphics.DrawString("Meal Plan Report", New Font("Arial", 16, FontStyle.Bold), brush, leftMargin, yPos)
        yPos += 40

        ' Check if there is data to print
        If mealPlanData.Rows.Count = 0 Then
            e.Graphics.DrawString("No data available.", font, brush, leftMargin, yPos)
            Exit Sub
        End If

        ' Print filtered meal plan data
        For Each row As DataRow In mealPlanData.Rows
            e.Graphics.DrawString("MealPlanID: " & row("mealPlanID").ToString(), font, brush, leftMargin, yPos)
            yPos += 30
            e.Graphics.DrawString("Start Date: " & Convert.ToDateTime(row("StartDate")).ToShortDateString(), font, brush, leftMargin, yPos)
            yPos += 30
            e.Graphics.DrawString("End Date: " & Convert.ToDateTime(row("EndDate")).ToShortDateString(), font, brush, leftMargin, yPos)
            yPos += 30
            e.Graphics.DrawString("NameOfMeal: " & row("NameOfMeal").ToString(), font, brush, leftMargin, yPos)
            yPos += 30
            e.Graphics.DrawString("Meals: " & row("Meals").ToString(), font, brush, leftMargin, yPos) ' Print Frequency
            yPos += 30
            e.Graphics.DrawString("totalcalories: " & row("totalcalories").ToString(), font, brush, leftMargin, yPos)
            yPos += 40
            e.Graphics.DrawString("Description: " & row("Description").ToString(), font, brush, leftMargin, yPos)
            yPos += 40
        Next

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub combobox4_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub
    'Public Function SSuggestMeals() As List(Of String)
    '    Dim suggestedMeals As New List(Of String)

    '    Try
    '        conn.Open()

    '        ' Get all meal recipes
    '        Dim mealQuery As String = "SELECT MealName, Ingredients FROM MealRecipes"
    '        Dim mealCommand As New OleDb.OleDbCommand(mealQuery, conn)
    '        Dim mealReader As OleDb.OleDbDataReader = mealCommand.ExecuteReader()

    '        While mealReader.Read()
    '            Dim mealName As String = mealReader("MealName").ToString()
    '            Dim requiredIngredients As String() = mealReader("Ingredients").ToString().Split(",")

    '            Dim allIngredientsAvailable As Boolean = True

    '            ' Check if all required ingredients exist in GroceryInventory
    '            For Each ingredient In requiredIngredients
    '                Dim checkQuery As String = "SELECT COUNT(*) FROM GroceryItem WHERE ItemName=@Ingredient AND Quantity > 0"
    '                Dim checkCommand As New OleDb.OleDbCommand(checkQuery, conn)
    '                checkCommand.Parameters.AddWithValue("@Ingredients", ingredient.Trim())

    '                Dim ingredientCount As Integer = Convert.ToInt32(checkCommand.ExecuteScalar())

    '                If ingredientCount = 0 Then
    '                    allIngredientsAvailable = False
    '                    Exit For
    '                End If
    '            Next

    '            ' If all ingredients are available, add the meal to suggested list
    '            If allIngredientsAvailable Then
    '                suggestedMeals.Add(mealName)
    '            End If
    '        End While

    '        mealReader.Close()
    '    Catch ex As Exception
    '        MsgBox("Error suggesting meals: " & ex.Message, MsgBoxStyle.Critical, "Database Error")
    '    Finally
    '        conn.Close()
    '    End Try

    '    Return suggestedMeals
    'End Function
    Public Function SuggestMeals() As List(Of String)
        Dim suggestedMeals As New List(Of String)

        Try
            Using connect As New OleDbConnection(connectionString)
                connect.Open()

                ' Get all meal recipes
                Dim mealQuery As String = "SELECT MealName, Ingredients FROM MealRecipes"
                Dim mealCommand As New OleDb.OleDbCommand(mealQuery, connect)
                Dim mealReader As OleDb.OleDbDataReader = mealCommand.ExecuteReader()

                While mealReader.Read()
                    Dim mealName As String = mealReader("MealName").ToString()
                    Dim requiredIngredients As String() = mealReader("Ingredients").ToString().Split(",")

                    Dim allIngredientsAvailable As Boolean = True

                    ' Check if all required ingredients exist in GroceryInventory and are not expired
                    For Each ingredient In requiredIngredients
                        Dim trimmedIngredient As String = ingredient.Trim()
                        Dim checkQuery As String = "SELECT ExpiryDate FROM GroceryItem WHERE ItemName=@Ingredients AND Quantity > 0"
                        Dim checkCommand As New OleDb.OleDbCommand(checkQuery, connect)
                        checkCommand.Parameters.AddWithValue("@Ingredients", trimmedIngredient)

                        Dim expirationDate As Object = checkCommand.ExecuteScalar()

                        ' Check if the ingredient exists and its expiration date
                        If expirationDate Is Nothing Then
                            allIngredientsAvailable = False
                        Else
                            ' Validate that the ingredient is not expired
                            If Convert.ToDateTime(expirationDate) < DateTime.Now Then
                                allIngredientsAvailable = False
                            End If
                        End If

                        If Not allIngredientsAvailable Then
                            Exit For
                        End If
                    Next

                    ' If all ingredients are available and not expired, add the meal to suggested list
                    If allIngredientsAvailable Then
                        suggestedMeals.Add(mealName)
                    End If
                End While
                mealReader.Close()
                'EndUsing
            End Using

        Catch ex As Exception
            MsgBox("Error suggesting meals: " & ex.Message, MsgBoxStyle.Critical, "Database Error")
        Finally
            conn.Close()
        End Try

        Return suggestedMeals
    End Function


    'Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

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

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click


        If lstSuggestedMeals.SelectedItem Is Nothing Then
            MsgBox("Please select a meal to save.", MsgBoxStyle.Exclamation, "Select Meal")
            Exit Sub
        End If

        Dim selectedMeal As String = lstSuggestedMeals.SelectedItem.ToString()

        Try
            conn.Open()

            Dim query As String = "INSERT INTO MealPlan (NameOfMeal, StartDate, EndDate) VALUES (@MealName, @StartDate, @EndDate)"
            Dim command As New OleDb.OleDbCommand(query, conn)

            command.Parameters.AddWithValue("@NameOfMeal", selectedMeal)
            command.Parameters.AddWithValue("@StartDate", Date.Today)
            command.Parameters.AddWithValue("@EndDate", Date.Today.AddDays(1)) ' One-day plan

            command.ExecuteNonQuery()
            MsgBox("Meal Plan Saved!", MsgBoxStyle.Information, "Success")

        Catch ex As Exception
            MsgBox("Error saving meal plan: " & ex.Message, MsgBoxStyle.Critical, "Database Error")
        Finally
        End Try
    End Sub
    Private Sub lstSuggestedMeals_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstSuggestedMeals.SelectedIndexChanged
        If lstSuggestedMeals.SelectedIndex <> -1 Then
            lstSuggestedMeals.Text = lstSuggestedMeals.SelectedItem.ToString
        End If
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click


        'Module1.Mains()

        lstSuggestedMeals.Items.Clear()
        TextBox2.ReadOnly = True
        Dim meals As List(Of String) = SuggestMeals()

        If meals.Count > 0 Then

            For Each meal In meals

                lstSuggestedMeals.Items.Add(meals)

            Next
            MsgBox("Meal Suggestions have been prepared with current Greocery Items.", MsgBoxStyle.Information, "No Available Meals")
            'FetchAlternativeMeals(SuggestMeals)
        Else
            MsgBox("No meals can be prepared with current inventory.", MsgBoxStyle.Exclamation, "No Available Meals")
        End If
    End Sub


End Class






