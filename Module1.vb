Imports System.IO
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Data
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net.Http
Imports System.Net

Module Module1
    Public personal As New List(Of person)
    Public conn As New OleDbConnection(Module1.connectionString)
    Public Const connectionString As String = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Austin\Documents\Visual Studio 2010\Projects\My hello\HOUSEHOLD MANAGEMENT SYSTEM.vb1\HOUSEHOLD MANAGEMENT SYSTEM.vb1\Austin.accdb"

    'Public Sub LoadPersonDataFromFile(filePath As String)
    '    If File.Exists(filePath) Then
    '        Try
    '            Using reader As New StreamReader(filePath)
    '                Dim person As person = Nothing
    '                While Not reader.EndOfStream
    '                    Dim line As String = reader.ReadLine()
    '                    If line.StartsWith("Name: ") Then
    '                        If person IsNot Nothing Then
    '                            personal.Add(person) ' Add the previous person to the list first  
    '                        End If
    '                        person = New person()
    '                        person.FirstName = line.Substring("Name: ".Length).Trim()
    '                    ElseIf line.StartsWith("Surname: ") Then
    '                        person.Lastname = line.Substring("Surname: ".Length).Trim()
    '                    ElseIf line.StartsWith("Contact Number: ") Then
    '                        person.ContactNumber = Integer.Parse(line.Substring("Contact Number: ".Length).Trim())
    '                    ElseIf line.StartsWith("eMail: ") Then
    '                        person.Email = line.Substring("eMail: ".Length).Trim()
    '                    ElseIf line.StartsWith("Role: ") Then
    '                        person.Role = line.Substring("Role: ".Length).Trim()
    '                    ElseIf line.StartsWith("Interest: ") Then
    '                        person.Interest = line.Substring("Interest: ".Length).Trim()
    '                    ElseIf line.StartsWith("Date Of Birth: ") Then
    '                        person.DateOfBirth = line.Substring("Date Of Birth: ".Length).Trim()
    '                    ElseIf line.StartsWith("Gender: ") Then
    '                        person.Gender = line.Substring("Gender: ".Length).Trim()
    '                    End If
    '                End While

    '                ' Add the last person to the list if we reached the end of the file  

    '                If person IsNot Nothing Then
    '                    personal.Add(person)
    '                End If
    '            End Using
    '        Catch ex As Exception
    '            Windows.Forms.MessageBox.Show("Error loading personnel data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        End Try
    '    Else
    '        MessageBox.Show("Personnel file not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End If
    ''    End 
    'End Sub
    Public Function GetPersonnelList() As List(Of person)
        Return personal
    End Function

    Public Sub BtnSaved()

        Try
            Debug.WriteLine("Entering btnSubmit")
            Dim Iventry As New Inventory With {
           .ItemID = Inventory.TextBox1.Text,
            .ItemName = Inventory.TextBox2.Text,
           .Description = Inventory.TextBox3.Text,
           .Quantity = Inventory.TextBox4.Text,
            .Unit = Inventory.TextBox5.Text,
           .Category = Inventory.TextBox6.Text,
           .ReorderLevel = Inventory.TextBox7.Text,
           .PricePerItem = Inventory.TextBox8.Text,
           .DateAdded = Inventory.DateTimePicker1.Text}


            Using connect As New OleDbConnection(Module1.connectionString)
                connect.Open()

                Dim tablename As String = "Inventory"

                Dim cmd As New OleDbCommand("INSERT INTO Inventory ([ItemID],[Names],[Description],[Quantity],[Unit],[Category],[ReorderLevel],[priceperItem],[DateAdded]) VALUES (@ItemID ,@Names ,@Description, @Quantity, @Unit, @Category, @ReorderLevel, @priceperunit, @DateAdded)", Module1.conn)
                'set the parameter values from the UI controls
                'class declaration  


                'params

                cmd.Parameters.AddWithValue("ItemID", Iventry.ItemID)
                cmd.Parameters.AddWithValue("ItemName", Iventry.ItemName)
                cmd.Parameters.AddWithValue("Description", Iventry.Description)
                cmd.Parameters.AddWithValue("Quantity", Iventry.Quantity)
                cmd.Parameters.AddWithValue("Unit", Iventry.Unit)
                cmd.Parameters.AddWithValue("Category", Iventry.Category)
                cmd.Parameters.AddWithValue("ReorderLevel", Iventry.ReorderLevel)
                cmd.Parameters.AddWithValue("PriceperUnit", Iventry.PricePerItem)
                cmd.Parameters.AddWithValue("DateAdded", Iventry.DateAdded)

                'cmd.ExecuteNonQuery()

            End Using



            MsgBox("Expense Inventoryitems Added!" & vbCrLf &
                 "ItemID: " & Iventry.ItemID & vbCrLf &
                   "ItemName: " & Iventry.ItemName.ToString() & vbCrLf &
                   "Description: " & Iventry.Description & vbCrLf &
                   "Quantity: " & Iventry.Quantity & vbCrLf &
                   "Unit: " & Iventry.Unit & vbCrLf &
                   "category: " & Iventry.Category & vbCrLf &
                   "ReorderLevel: " & Iventry.ReorderLevel & vbCrLf &
                   "PricePerUnit: " & Iventry.PricePerItem & vbCrLf &
                   "DateAdded: " & Iventry.DateAdded, vbInformation, "inventory Confirmation")




        Catch ex As OleDbException
            Debug.WriteLine($"Database error: {ex.Message}")
            MessageBox.Show("Error saving test to database. Please check the connection.", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Debug.WriteLine($"General error: {ex.Message}")
            'MessageBox.Show("An unexpected error occurred.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            conn.Close()
        Finally

        End Try
    End Sub
    Public Sub BtnEditInventory()

        If Inventory.DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Debug.WriteLine("Entering btnEdit")
            Dim ItemID As Integer = Inventory.TextBox1.Text
            Dim ItemName As String = Inventory.TextBox2.Text
            Dim Description As String = Inventory.TextBox3.Text
            Dim Quantity As Integer = Inventory.TextBox4.Text
            Dim Unit As String = Inventory.TextBox5.Text
            Dim Category As String = Inventory.TextBox6.Text
            Dim ReorderLevel As String = Inventory.TextBox7.Text
            Dim PriceperItem As Integer = Inventory.TextBox8.Text
            Dim DateAdded As DateTime = Inventory.DateTimePicker1.Text


            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()

                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = Inventory.DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Create an OleDbCommand to update the expense data in the database  

                Dim cmd As New OleDbCommand("UPDATE [Inventory] SET [ItemID] = ?, [ItemName]  = ?, [Description] = ?, [Quantity] = ?, [Unit] = ?, [Category] = ?,[ReorderLevel]=?,[PriceperItem]=?[ DateAdded]
WHERE [ID] = ?", conn)
                ' Set the parameter values from the UI controls  

                cmd.Parameters.AddWithValue("@ItemID", Inventory.TextBox1.Text)
                cmd.Parameters.AddWithValue("@ItemName", Inventory.TextBox2.Text)
                cmd.Parameters.AddWithValue("@Description", Inventory.TextBox3.Text)
                cmd.Parameters.AddWithValue("@Quantity", Inventory.TextBox4.Text)
                cmd.Parameters.AddWithValue("@Unit", Inventory.TextBox5.Text)
                cmd.Parameters.AddWithValue("@Category", Inventory.TextBox6.Text)
                cmd.Parameters.AddWithValue("@ReorderLevel", Inventory.TextBox7.Text)
                cmd.Parameters.AddWithValue("@ PriceperItem", Inventory.TextBox8)
                cmd.Parameters.AddWithValue("@DateAdded", Inventory.DateTimePicker1.Text)

                'cmd.Parameters.AddWithValue("@ID", ExpenseID) ' Primary key for matching record  
                'cmd.ExecuteNonQuery()


                MsgBox("Inventory Updated Successfuly!", vbInformation, "Update Confirmation")


            End Using
        Catch ex As OleDbException
            MessageBox.Show($"Error updating Expenses in database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        conn.Close()
    End Sub

    Public Sub btnDelete()


        ' Check if there are any selected rows in the DataGridView for expenses  
        If Inventory.DataGridView1.SelectedRows.Count > 0 Then
            ' Get the selected row  
            Dim selectedRow As DataGridViewRow = Inventory.DataGridView1.SelectedRows(0)

            ' Assuming there is an "ID" column which is the primary key in the database  
            Dim InventoryId As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  

            ' Confirm deletion  
            Dim confirmationResult As DialogResult = MessageBox.Show("Are you sure you want to delete this expense?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If confirmationResult = DialogResult.Yes Then
                ' Proceed with deletion  
                Try
                    Using conn As New OleDbConnection(Module1.connectionString)
                        conn.Open()

                        ' Create the delete command  
                        Dim cmd As New OleDbCommand("DELETE FROM [Inventory] WHERE [ID] = ?", conn)
                        cmd.Parameters.AddWithValue("@ID", InventoryId) ' Primary key for matching record  

                        ' Execute the delete command  
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show("Inventory deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
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
            conn.Close()
        End If
    End Sub

    Public Sub FilterChores(Frequency As String, Priority As String)
        Dim Tasktable As DataTable()

        Try

            conn.Open()
            Dim query As String = "SELECT *FROM CHORES WHERE 1=1"
            If Not String.IsNullOrEmpty(Frequency) Then
                query &= "AND Frequency =@Frequency"
            End If

            If Not String.IsNullOrEmpty(Priority) Then
                query &= "AND Priority =@Priority"
            End If
            Dim command As New OleDb.OleDbCommand(query, conn)
            If Not String.IsNullOrEmpty(Frequency) Then
                command.Parameters.AddWithValue("@Frequency", Frequency)
            End If
            If Not String.IsNullOrEmpty(Priority) Then
                command.Parameters.AddWithValue("@Priority", Priority)

            End If

            Dim Adapter As New OleDb.OleDbDataAdapter(command)
            'Adapter.Fill(Tasktable)
            Chores.DataGridView1.DataSource = Tasktable

        Catch ex As Exception

            MsgBox("Error Filtering tasks " & ex.Message, MsgBoxStyle.Critical, "DataBase error")
            Debug.WriteLine($"stack Trace:{ex.StackTrace}")
        Finally
            conn.Close()
        End Try
    End Sub



    '    Private Sub LoadExpensesDataFromFile(filePath As String)
    '        If File.Exists(filePath) Then
    '            Try
    '                Using reader As New StreamReader(filePath)
    '                    Dim Expenses As Expenses = Nothing
    '                    While Not reader.EndOfStream
    '                        Dim line As String = reader.ReadLine()
    '                        If line.StartsWith("Name: ") Then
    '                            If Expenses IsNot Nothing Then
    '                                Expense.Add(Expenseses) ' Add the previous person to the list first  
    '                            End If
    '                            expense = New Expenses()
    '                            Expenses.Dates = line.Substring("Dates: ".Length).Trim()
    '                        ElseIf line.StartsWith("amount: ") Then
    '                            Expenses.Amount = line.Substring("amount: ".Length).Trim()
    '                        ElseIf line.StartsWith("Currency ") Then
    '                            Expenses.frequency = Integer.Parse(line.Substring("Frequency: ".Length).Trim())
    '                        ElseIf line.StartsWith("ApprovalStatus: ") Then
    '                            Expenses.approvalStatus = line.Substring("approvalstatus: ".Length).Trim()
    '                        ElseIf line.StartsWith("tags: ") Then
    '                            Expenses.Tags = line.Substring("Tags: ".Length).Trim()
    '                        ElseIf line.StartsWith("paymentmethod: ") Then
    '                            Expenses.payment = line.Substring("paymentmethod: ".Length).Trim()

    '                        End If
    '                    End While

    '                    Add the last person to the list if we reached the end of the file  

    '                    If person IsNot Nothing Then
    '                        expense.Add(Expenses)
    '                    End If
    '                End Using
    '            Catch ex As Exception
    '                MessageBox.Show("Error loading personnel data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            End Try
    '        Else
    '            MessageBox.Show("Personnel file not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)








    Sub Main()
            ' Database connection string (Update with your actual database details)
            Dim connectionString As String = "Server=YOUR_SERVER;Database=YOUR_DATABASE;User Id=YOUR_USERNAME;Password=YOUR_PASSWORD;"

            ' Query to check items with zero quantity
            Dim query As String = "SELECT ItemName FROM Inventory WHERE Quantity <= 0"

            Try
                Using connection As New SqlConnection(connectionString)
                    connection.Open()

                    Using command As New SqlCommand(query, connection)
                        Using reader As SqlDataReader = command.ExecuteReader()
                            If reader.HasRows Then
                                Dim outOfStockItems As String = "Out of Stock Items:" & vbCrLf

                                While reader.Read()
                                    outOfStockItems &= reader("ItemName").ToString() & vbCrLf
                                End While

                                ' Send email notification
                                SendEmail(outOfStockItems)
                            End If
                        End Using
                    End Using
                End Using
            Catch ex As Exception
                Console.WriteLine("Error: " & ex.Message)
            End Try
        End Sub
    'Public Function SuggestMeals() As List(Of String)
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
    '                checkCommand.Parameters.AddWithValue("@Ingredient", ingredient.Trim())

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
    Public Function IsInternetAvailable() As Boolean

        Try

            Using client As New WebClient()

                Using stream = client.OpenRead("http://www.google.com")

                    Return True

                End Using

            End Using

        Catch ex As Exception

            Return False

        End Try

    End Function
    Public Sub SendEmail(body As String)
        Try
            Dim smtpClient As New SmtpClient("smtp.yourmailserver.com") ' Update SMTP server
            smtpClient.Port = 587 ' Adjust if needed (e.g., 465 for SSL, 587 for TLS)
            smtpClient.Credentials = New Net.NetworkCredential("your-email@example.com", "your-password")
            smtpClient.EnableSsl = True

            Dim mail As New MailMessage()
            mail.From = New MailAddress("your-email@example.com")
            mail.To.Add("recipient@example.com") ' Update recipient email
            mail.Subject = "Stock Alert: Items Out of Stock"
            mail.Body = body

            smtpClient.Send(mail)
            Console.WriteLine("Email sent successfully!")
        Catch ex As Exception
            Console.WriteLine("Email sending failed: " & ex.Message)
        End Try
    End Sub
End Module






