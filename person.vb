Imports System.IO
Imports System.Windows.Forms
Imports System.Data.OleDb
Public Class person
    Dim connection As New OleDb.OleDbConnection(Module1.connectionString)
    Private personal As New List(Of person)
    Private Sub btnSubmit_Click(sender As System.Object, e As System.EventArgs) Handles btnSubmit.Click
        'Dim str As String
        'Create Instance of person Information 

        Try

            Dim person As New person()

            ' Assign values from textBoxes to the   PersonInfomation  
            person.FirstName = TextBox1.Text
            person.Lastname = TextBox2.Text
            person.DateOfBirth = DateTimePicker1.Value.ToShortDateString()
            person.Gender = ComboBox1.Text
            person.ContactNumber = TextBox4.Text
            person.Email = TextBox5.Text
            person.Role = ComboBox2.Text
            person.MaritalStatus = ComboBox3.Text
            person.Interest = TextBox3.Text


            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()

                Dim tableName As String = "personaldetails"

                ' Create an OleDbCommand to insert the personn data into the database  
                Dim cmd As New OleDbCommand($"INSERT INTO [personaldetails] ([FirstName],[Lastname],[DateOfbirth],[Gender],[ContactNumber],[MaritalStatus],[Role],[Interest]) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", conn)




                ' Set the parameter values from the UI controls  
                For Each persons As person In personal
                    cmd.Parameters.Clear()
                    cmd.Parameters.AddWithValue("@FirstName", persons.FirstName)
                    cmd.Parameters.AddWithValue("@Last name", persons.Lastname)
                    cmd.Parameters.AddWithValue("@DateOfbirth", persons.DateOfBirth)
                    cmd.Parameters.AddWithValue("@Gender", persons.Gender)
                    cmd.Parameters.AddWithValue("@ContactNumber", persons.ContactNumber)
                    cmd.Parameters.AddWithValue("@Email", persons.Email)
                    cmd.Parameters.AddWithValue("@MaritalStatus", persons.MaritalStatus)
                    cmd.Parameters.AddWithValue("@Role", persons.Role)
                    cmd.Parameters.AddWithValue("@Interest", persons.Interest)
                    conn.Close()
                    ' Execute the SQL command to insert the data  
                    cmd.ExecuteNonQuery()


                    ' Display a confirmation messageBox  
                    MsgBox("personal Information added!" & vbCrLf & "FirstName:" & person.FirstName & vbCrLf &
                    "Lastname:" & person.Lastname & vbCrLf &
                    "BirthDate:" & person.DateOfBirth & vbCrLf &
                    "Gender:" & person.Gender & vbCrLf &
                    "ContactNumber:" & person.ContactNumber & vbCrLf &
                    "Email:" & person.Email & vbCrLf &
                    "Role:" & person.Role & vbCrLf &
                    "Maritalstatus:" & person.MaritalStatus & vbCrLf &
                    "Interest:" & person.Interest, vbInformation, "Personal confirmation")


                Next
            End Using
        Catch ex As OleDbException
            MessageBox.Show("Error saving personaldetails to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("Unexpected Error: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    Public Sub loadPersonaldataFromDatabase()
        Try
            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()
                'Dim tableName As String = "personaldetails"
                Dim cmd As New OleDbCommand("SELECT * FROM personaldetails", conn)
                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable()
                da.Fill(dt)

                DataGridView1.DataSource = dt

            End Using

        Catch ex As Exception
            'MessageBox.Show("Error Loading personnel data from database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show("Error Loading personnel data from database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Set tooltips For buttons
        ToolTip1.SetToolTip(Button4, "Dashboard")
        ToolTip1.SetToolTip(btnSubmit, "Save")
        ToolTip1.SetToolTip(BtnEdit, "Edit")
        ToolTip1.SetToolTip(btnClear, "Clear")
        ToolTip1.SetToolTip(btnGetInput, "Input")
        ToolTip1.SetToolTip(Button1, "Refresh")
        ToolTip1.SetToolTip(Button2, "Filter")
        ToolTip1.SetToolTip(Button3, "Delete")

    End Sub

    Private Sub populatedDataGridview()
        'clear Existing rows
        DataGridView1.Rows.Clear()
        'Add each expense to the GridView
        For Each per As person In personal
            DataGridView1.Rows.Add(per.FirstName, per.Lastname, per.DateOfBirth, per.Gender, per.ContactNumber, per.Email, per.Role, per.MaritalStatus, per.Interest)

        Next
    End Sub

    'Private personal  As New List(Of Expenses)
    Private Sub PopulateComboBox()
        ComboBox2.Items.Clear()
        For Each persons As person In personal
            ComboBox2.Items.Add(persons.FirstName & " " & persons.Lastname)
        Next
    End Sub

    Private Sub personal_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Dim connections As New OleDb.OleDbConnection(connectionString)
        'PopulateComboBox()

        conn.Open()

        Try
            ' Display the connection status on a button with a green background  
            personaldetails.Text = "Connected"
            personaldetails.BackColor = System.Drawing.Color.Green
            personaldetails.ForeColor = System.Drawing.Color.White
        Catch ex As Exception
            ' Display the connection status on a button with a red background  
            personaldetails.Text = "Not Connected"
            personaldetails.BackColor = System.Drawing.Color.Red
            personaldetails.ForeColor = System.Drawing.Color.White

            ' Display an error message  
            MessageBox.Show("Error connecting to the database" & ex.Message, "Database Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            loadPersonaldataFromDatabase()

            ' Close the database connection  
            conn.Close()
        End Try

        ' Disable certain buttons if the connection is not established  
        Button3.Enabled = personaldetails.Text = "Connected"
        Button4.Enabled = personaldetails.Text = "Connected"


    End Sub

    Private Sub btnGetInput_Click(sender As System.Object, e As System.EventArgs) Handles btnGetInput.Click
        TextBox1.Text = InputBox("Enter your Name")
        TextBox2.Text = InputBox("Enter your lastName")
        ComboBox1.Text = InputBox("Enter your gender")
        TextBox3.Text = InputBox("Enter your Interest")
        TextBox4.Text = InputBox("Enter your ContactNumber")
        ComboBox2.Text = InputBox("Enter your Role")
        ComboBox3.Text = InputBox("Enter your MaritalStatus")

    End Sub

    'Public Sub LoadPersonnelDataFromDatabase()

    '    Try
    '        Using conn As New OleDbConnection(Module1.connectionString)
    '            conn.Open()

    '            ' Update the table name if necessary  
    '            Dim tableName As String = "PersonalDetails"

    '            ' Create an OleDbCommand to select the data from the database  
    '            Dim cmd As New OleDbCommand("SELECT * FROM {tableName}", conn)

    '            ' Create a DataAdapter and fill a DataTable  
    '            Dim da As New OleDbDataAdapter(cmd)
    '            Dim dt As New DataTable()
    '            da.Fill(dt)

    '            ' Bind the DataTable to the DataGridView  
    '            DataGridView1.DataSource = dt
    '        End Using
    '    Catch ex As OleDbException
    '        MessageBox.Show("Error loading personnel data from database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    Catch ex As Exception
    '        MessageBox.Show("Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub

    'Private Sub saveToFile(ByVal person As person)

    '    Try
    '        'define the file path where the details will be saved

    '        Dim filePath As String = "C:\Users\user\Documents\Visual Studio 2010\Projects\My hello\HOUSEHOLD MANAGEMENT SYSTEM.vb1\HOUSEHOLD MANAGEMENT SYSTEM.vb1\Storage\Person.txt"

    '        'Dim filePath As String = "C:\Temp\personnel.txt"
    '        'create or open the file and append the details
    '        Using writer As New StreamWriter(filePath, True)
    '            writer.WriteLine("Name: " & person.FirstName)
    '            writer.WriteLine("Surname: " & person.Lastname)
    '            writer.WriteLine("Contact Number: " & person.ContactNumber)
    '            writer.WriteLine("eMail: " & person.Email)
    '            writer.WriteLine("Role: " & person.Role)
    '            writer.WriteLine("Interest: " & person.Interest)
    '            writer.WriteLine("Date Of Birth: " & person.DateOfBirth)
    '            writer.WriteLine("Gender: " & person.Gender)
    '            writer.WriteLine("..................")


    '        End Using
    '        MessageBox.Show("Error saving control texts: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

    '        MessageBox.Show("control Texts saved to file successfuly!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

    '    Finally
    '    End Try





    'Private Sub LoadPersonDataFromFile(filePath As String)
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
    '                    ElseIf line.StartsWith("last name: ") Then
    '                        person.Lastname = line.Substring("Surname: ".Length).Trim()
    '                    ElseIf line.StartsWith("DateOfBirth: ") Then
    '                        person.ContactNumber = Integer.Parse(line.Substring("Contact Number: ".Length).Trim())
    '                    ElseIf line.StartsWith("eMail: ") Then
    '                        person.Email = line.Substring("eMail: ".Length).Trim()
    '                    ElseIf line.StartsWith("Role: ") Then
    '                        person.Role = line.Substring("Role: ".Length).Trim()
    '                    ElseIf line.StartsWith("MaritalStatus: ") Then
    '                        person.Interest = line.Substring("Interest: ".Length).Trim()
    '                    ElseIf line.StartsWith("Date Of Birth: ") Then
    '                        person.DateOfBirth.Substring("birthdate".Length).Trim()
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
    '            MessageBox.Show("Error loading personnel data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        End Try
    '    Else
    '        MessageBox.Show("Personnel file not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End If
    'End Sub



    Private Sub btnClear_Click(sender As System.Object, e As System.EventArgs) Handles btnClear.Click
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()

    End Sub
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
            'Load the data from  the selected row into UI controls 
            'TextBox1.Text = selectedRow.Cells("FirstName").Value.ToString()
            'TextBox2.Text = selectedRow.Cells("LastName").Value.ToString()
            'DateTimePicker1.Text = selectedRow.Cells("Dates").Value.ToString()
            'ComboBox1.Text = selectedRow.Cells("Gender").Value.ToString()
            TextBox4.Text = selectedRow.Cells("ContactNumber").Value.ToString()
            'TextBox5.Text = selectedRow.Cells("Email").Value.ToString()
            ComboBox2.Text = selectedRow.Cells("Role").Value.ToString()
            ComboBox3.Text = selectedRow.Cells("MaritalStatus").Value.ToString()
            TextBox3.Text = selectedRow.Cells("Interest").Value.ToString()

        End If
    End Sub


    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs)
        GroceryItemvb.ShowDialog()

    End Sub

    Private Sub Button4_Click_1(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        DashBoard.Show()
        Me.Close()
    End Sub




    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles personaldetails.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles BtnEdit.Click
        If DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Dim FirstName As String = TextBox1.Text
            Dim LastName As String = TextBox2.Text
            Dim DateOfBirth As String = DateTimePicker1.Text ' Ensure this is of DateTime type  
            Dim Gender As String = ComboBox1.Text
            Dim ContactNumber As String = TextBox4.Text
            Dim Email As String = TextBox5.Text
            Dim Role As String = ComboBox2.Text
            Dim MaritalStatus As String = ComboBox3.Text
            Dim Interest As String = TextBox3.Text


            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()

                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                'Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                '' Create an OleDbCommand to update the expense data in the database  

                Dim cmd As New OleDbCommand("UPDATE personaldetails SET [FirstName] = ?, [LastName]  = ?, [DateOfBirth] = ?, [Gender] = ?, [ContactNumber] = ?, [MaritalStatus] = ?, [Interest] = ? WHERE [ID] = ?", conn)
                ' Set the parameter values from the UI controls  

                cmd.Parameters.AddWithValue("@FirstName", TextBox1.Text)
                cmd.Parameters.AddWithValue("@LastName", TextBox2.Text)
                cmd.Parameters.AddWithValue("@DateOfBirth", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@Gender", ComboBox1.Text)
                cmd.Parameters.AddWithValue("@ContactNumber", TextBox4.Text)
                cmd.Parameters.AddWithValue("@Email", TextBox5.Text)
                cmd.Parameters.AddWithValue("@Role", ComboBox2.SelectedItem.ToString)
                cmd.Parameters.AddWithValue("@ MaritalStatus", ComboBox3.SelectedItem.ToString)
                cmd.Parameters.AddWithValue("@Interest", TextBox3.Text)
                'cmd.Parameters.AddWithValue("@ID", personsID) ' Primary key for matching record  
                'cmd.ExecuteNonQuery()



                MsgBox("Expense Updated Successfuly!", vbInformation, "Update Confirmation")


                loadPersonaldataFromDatabase()
                'person.ClearControls(Me)

            End Using
        Catch ex As OleDbException
            MessageBox.Show($"Error updating Expenses in database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        ' Check if there are any selected rows in the DataGridView for expenses  
        If DataGridView1.SelectedRows.Count > 0 Then
            ' Get the selected row  
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Assuming there is an "ID" column which is the primary key in the database  
            Dim personaldetailsId As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  

            ' Confirm deletion  
            Dim confirmationResult As DialogResult = MessageBox.Show("Are you sure you want to delete this expense?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If confirmationResult = DialogResult.Yes Then
                ' Proceed with deletion  
                Try
                    Using conn As New OleDbConnection(Module1.connectionString)
                        conn.Open()

                        ' Create the delete command  
                        Dim cmd As New OleDbCommand("DELETE FROM [personaldetails] WHERE [ID] = ?", conn)
                        cmd.Parameters.AddWithValue("@ID", personaldetailsId) ' Primary key for matching record  

                        ' Execute the delete command  
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show("Expense deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ' Optionally refresh DataGridView or reload from database  
                            'PopulateDataGridView()
                        Else
                            MessageBox.Show("No  personaldetails  was deleted. Please check if the ID exists.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End Using

                Catch ex As Exception
                    MessageBox.Show($"An error occurred while deleting the personaldetails: {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            MessageBox.Show("Please select an personaldetails to delete.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub
End Class
