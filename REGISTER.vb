Imports System.Data.OleDb
Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO
Imports System.Data.SqlClient


Public Class regis
    Dim conn As New OleDb.OleDbConnection(Module1.connectionString)
    Dim txtFirstName As New TextBox
    Dim txtLastname As New TextBox
    Dim txtUsername As New TextBox
    Dim txtPassword As New TextBox
    Dim txtEmail As New TextBox
    Dim TxtRole As New TextBox
    Dim REGISTER As Object

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        DashBoard.Show()
        Me.Close()
    End Sub
    Private Sub REGISTER_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TxtFirstName.Focus()



        ' Set tooltips For buttons
        ToolTip1.SetToolTip(Button1, "Dashboard")
        ToolTip1.SetToolTip(btnLogin, "Log in")
        ToolTip1.SetToolTip(Btnupdate, "Update")
        ToolTip1.SetToolTip(BtnCancel, "Cancel")
        ToolTip1.SetToolTip(Button2, "Filter")

        'LadREGISTERDataFromDatabase()

        Dim connection As New OleDbConnection(connectionString)
        'Cheak Database connectivity 
        Try

            Debug.WriteLine("register load has been initialised")


            Dim lblFirstNames As New Label()
            lblFirstNames.Text = "FirstName"
            lblFirstNames.Name = "lblFirstName"
            lblFirstNames.Location = New Point(15, 50)
            lblFirstNames.AutoSize = True
            lblFirstNames.ForeColor = Color.Black
            Me.Controls.Add(lblFirstNames)
            Debug.WriteLine("Expense_ load: DateTime label added. ")

            Dim lblLastName As New Label()
            lblLastName.Text = "LastName"
            lblLastName.Name = "lblLastName"
            lblLastName.Location = New Point(15, 100)
            lblLastName.AutoSize = True
            Me.Controls.Add(lblLastName)
            Debug.WriteLine("Expense_ load: DateTime label added. ")

            Dim lblUserName As New Label()
            lblUserName.Text = "UserName"
            lblUserName.Name = "lblUserName"
            lblUserName.Location = New Point(15, 150)
            lblUserName.AutoSize = True
            Me.Controls.Add(lblUserName)
            Debug.WriteLine("Expense_ load: DateTime label added. ")

            Dim lblPassword As New Label()
            lblPassword.Text = "Password"
            lblPassword.Name = "lblPassword"
            lblPassword.Location = New Point(15, 200)
            lblPassword.AutoSize = True
            Me.Controls.Add(lblPassword)
            Debug.WriteLine("Expense_ load: DateTime label added. ")

            Dim lblEmail As New Label()
            lblEmail.Text = "Email"
            lblEmail.Name = "lblEmail"
            lblEmail.Location = New Point(15, 250)
            lblEmail.AutoSize = True
            Me.Controls.Add(lblEmail)
            Debug.WriteLine("Expense_ load: DateTime label added. ")

            Dim lblRole As New Label()
            lblRole.Text = "Role"
            lblRole.Name = "lblRole"
            lblRole.Location = New Point(15, 300)
            lblRole.AutoSize = True
            Me.Controls.Add(lblRole)
            Debug.WriteLine("Expense_ load: DateTime label added. ")

            ''Dim txtFirstName As New TextBox()
            'txtFirstName.Name = "txtFirstName"
            txtFirstName.Location = New Point(320, 50)
            txtFirstName.AutoSize = True
            Me.Controls.Add(txtFirstName)
            Debug.WriteLine("Expense_ load: DateTime textbox added. ")

            'Dim txtLastName As New TextBox()
            'txtLastName.Name = "txtFirstName"
            txtLastname.Location = New Point(320, 100)
            txtLastname.AutoSize = True
            Me.Controls.Add(txtLastname)
            Debug.WriteLine("Expense_ load: DateTime textbox added. ")

            'Dim txtUserName As New TextBox()
            'txtUserName.Name = "txtUserName"
            txtUsername.Location = New Point(320, 150)
            txtUsername.AutoSize = True
            Me.Controls.Add(txtUsername)
            Debug.WriteLine("Expense_ load: DateTime textbox added. ")

            'Dim txtpassword As New TextBox()
            'txtpassword.Name = "txtpassword"
            txtPassword.Location = New Point(320, 200)
            txtPassword.AutoSize = True
            Me.Controls.Add(txtPassword)
            Debug.WriteLine("Expense_ load: DateTime textbox added. ")

            'Dim txtEmail As New TextBox()
            'txtEmail.Name = "txtEmail"
            txtEmail.Location = New Point(320, 250)
            txtEmail.AutoSize = True
            Me.Controls.Add(txtEmail)
            Debug.WriteLine("Expense_ load: DateTime textbox added. ")


            'Dim txtRole As New TextBox()
            'txtRole.Name = "txtRole"
            TxtRole.Location = New Point(320, 300)
            TxtRole.AutoSize = True
            Me.Controls.Add(TxtRole)
            Debug.WriteLine("Expense_ load: DateTime textbox added. ")

            Dim btnCancel As New Button()
            btnCancel.Text = "cancel"
            btnCancel.Location = New Point(15, 350)
            btnCancel.AutoSize = True
            Me.Controls.Add(btnCancel)
            Debug.WriteLine("Expense_ load: DateTime Button added. ")

            Dim btnRegister As New Button()
            btnRegister.Text = "Register"
            btnRegister.Location = New Point(15, 400)
            btnRegister.AutoSize = True
            Me.Controls.Add(btnRegister)
            Debug.WriteLine("Expense_ load: DateTime Button added. ")

        Catch ex As Exception
            Debug.WriteLine("Form failed to load successfully")
            Debug.WriteLine($"Error in expense_load when creating  DateTime label:{ex.Message}")
            Debug.WriteLine($"Stack Trace:{ex.StackTrace}")
            MessageBox.Show("An unexpexted error occured during Form load.", "Expensesload Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show("An error occured while displaying the date  and time. ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs)


        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            'load the data from  the selected row into UI controls 

            txtFirstName.Text = selectedRow.Cells("Firstname").Value.ToString()
            txtLastname.Text = selectedRow.Cells("LastName").Value.ToString()
            txtUsername.Text = selectedRow.Cells("UserName").Value.ToString()
            txtPassword.Text = selectedRow.Cells("password").Value.ToString()
            txtEmail.Text = selectedRow.Cells("Email").Value.ToString()
            TxtRole.Text = selectedRow.Cells("Role").Value.ToString()


            Debug.WriteLine("No row selected, exiting selectedchanged _Click.")
            MessageBox.Show("Please select an register to update.", "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
    End Sub
    Public Sub LoadREGISTERDataFromDatabase()
        Try
            Using connect As New OleDbConnection(Module1.connectionString)
                connect.Open()

                ' Update the table name if necessary  
                Dim tableName As String = "Logindetails"

                ' Create an OleDbCommand to select the data from the database  
                Dim cmd As New OleDbCommand($"SELECT * FROM {tableName}", connect)

                ' Create a DataAdapter and fill a DataTable  
                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable()
                da.Fill(dt)

                ' Bind the DataTable to the DataGridView  
                DataGridView1.DataSource = dt
            End Using
        Catch ex As OleDbException
            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            MessageBox.Show($"Error loading person data from database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub loadregisdataFromDatabase()
        Try
            Debug.WriteLine("populated connected succesfully")
            Using connect As New OleDbConnection(Module1.connectionString)
                connect.Open()

                Dim tableName As String = "Logindetails"
                Dim cmd As New OleDbCommand($"SELECT * FROM {tableName}", connect)

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
            For Each exp As regis In REGISTER
                DataGridView1.Rows.Add(exp.FirstName, exp.LastName, exp.Username, exp.password, exp.Email, exp.Role())
            Next
        Catch ex As Exception
            Debug.WriteLine("fail to populate")
            Debug.WriteLine($"Stack Trace:{ex.StackTrace}")
            'MessageBox.Show("An unexpexted error occured during DataGridview")
        End Try
    End Sub
    Private Sub Btnupdate_Click(sender As Object, e As EventArgs) Handles Btnupdate.Click

        Debug.WriteLine("Entering btnSubmit")

        Try

            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()
                Debug.WriteLine("User confirmed btnSubmit")


                ' Update the table name if necessary  
                'Dim tableName As String = "Personnel"

                ' Create an OleDbCommand to insert the person data into the database  
                Dim cmd As New OleDbCommand("INSERT INTO [Logindetails] ([FirstName], [LastName], [UserName], [Password], [Email], [Role]) VALUES (?, ?, ?, ?, ?, ?)", conn)

                ' Set the parameter values from the UI controls 
                'Class declaretions
                Dim Regis As New regis

                'Assign Values 
                Regis.FirstName = txtFirstName.Text
                Regis.LastName = txtLastname.Text
                Regis.Username = txtUsername.Text
                Regis.password = txtPassword.Text
                Regis.Email = txtEmail.Text
                Regis.Role = TxtRole.Text


                'For Each person As person In Personal
                cmd.Parameters.Clear()

                'cmd.Parameters.AddWithValue("@ID",TextBox8)
                cmd.Parameters.AddWithValue("@FirstName", Regis.FirstName)
                cmd.Parameters.AddWithValue("@LastName", Regis.LastName)
                cmd.Parameters.AddWithValue("@UserName", Regis.Username)
                cmd.Parameters.AddWithValue("@Password", Regis.password)
                cmd.Parameters.AddWithValue("@Email", Regis.Email)
                cmd.Parameters.AddWithValue("@Role", Regis.Role)

                MsgBox(" You are now added as a member !" & vbCrLf &
                      "FirstName: " & Regis.FirstName & vbCrLf &
                      "LastName: " & Regis.LastName & vbCrLf &
                      "ContactNumber:" & Regis.Username & vbCrLf &
                      "Interest: " & Regis.password & vbCrLf &
                      "ContactNumber:" & Regis.Email & vbCrLf &
                      "DateOfBirth: " & Regis.Role & vbCrLf, vbInformation, "Credentials  confirmation")

                MessageBox.Show("LoginDetails information saved to Database successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'DashBoard.ShowDialog()
                ' Execute the SQL command to insert the data  
                cmd.ExecuteNonQuery()
                'Next

            End Using
        Catch ex As OleDbException
            'MessageBox.Show("Error saving LoginDetails to database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            MessageBox.Show("Error saving LoginDetails to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            MessageBox.Show("Unexpected Error: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        conn.Close()
        'loadregisdataFromDatabase()
        LoadREGISTERDataFromDatabase()

        'Try
        '    Debug.WriteLine("Entering btnSubmit successfuly")

        '    Using connect As New OleDbConnection(Module1.connectionString)
        '        conn.Open()
        '        Dim tablename As String = "Logindetails"

        '        Dim Cmd As New OleDbCommand($"INSERT INTO {tablename} ([FirstName], [LastName], [Username], [password], [Email], [ Role])VALUES ( ?, ?, ?, ?, ?, ?,)", connect)


        '        ' Assign values  
        '        Dim regis As New regis()
        '        regis.txtFirstName = txtFirstName
        '        regis.txtLastname = txtLastname
        '        regis.txtUsername = txtUsername
        '        regis.txtPassword = txtPassword
        '        regis.txtEmail = txtEmail
        '        regis.TxtRole = TxtRole

        '        Cmd.Parameters.Clear()

        '        Cmd.Parameters.AddWithValue("@FirstName", regis.txtFirstName.Text)
        '        Cmd.Parameters.AddWithValue("@LastName", regis.txtLastname.Text())
        '        Cmd.Parameters.AddWithValue("@UserName", regis.txtUsername.Text())
        '        Cmd.Parameters.AddWithValue("@password", regis.txtPassword.Text)
        '        Cmd.Parameters.AddWithValue("@Email", regis.txtEmail.Text())
        '        Cmd.Parameters.AddWithValue("@Role", regis.TxtRole.Text())


        '        MsgBox("Register Information Added!" & vbCrLf &
        '            "FirstName: " & regis.txtFirstName.Text & vbCrLf &
        '            "Lastname: " & regis.txtLastname.Text & vbCrLf &
        '            "UserName: " & regis.txtUsername.Text & vbCrLf &
        '            "Password: " & regis.txtPassword.Text & vbCrLf &
        '            "Email: " & regis.txtEmail.Text & vbCrLf &
        '            "Role " & regis.TxtRole.Text, vbInformation, "Credentials Confirmation")

        '        conn.Close()
        '        populatedDatagridview()
        '        loadregisdataFromDatabase()



        '        MessageBox.Show("Logindetails  saved successfully to database!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        '        'Execute the SQL command to the insert data
        '        Cmd.ExecuteNonQuery()

        '    End Using



        'Catch ex As OleDbException
        '    MessageBox.Show("Unexpected Error occured: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    Debug.WriteLine($"Stack Trace:{ex.StackTrace}")
        'Catch ex As Exception
        '    Debug.WriteLine("Exiting btnSubmit successfully")
        'End Try


    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As EventArgs) Handles BtnCancel.Click
        Me.Close()
    End Sub
    'Public Class RegisterForm
    '    Private connectionString As String = "Data Source=YOUR_SERVER;Initial Catalog=YOUR_DATABASE;Integrated Security=True"

    '    ' Function to check if the username exists
    '    Private Function UsernameExists(username As String) As Boolean
    '        Dim query As String = "SELECT COUNT(*) FROM Users WHERE Username = @Username"
    '        Using con As New SqlConnection(connectionString)
    '            Using cmd As New SqlCommand(query, con)
    '                cmd.Parameters.AddWithValue("@Username", username)
    '                con.Open()
    '                Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
    '                Return count > 0
    '            End Using
    '        End Using
    '    End Function

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click




            conn.Open()
            Dim query As String = "SELECT COUNT(*) FROM Logindetails WHERE FirstName = ? AND LastName =?"
            Using cmd As New OleDbCommand(query, conn)

                cmd.Parameters.AddWithValue("@FirstName", txtUsername.Text)
                cmd.Parameters.AddWithValue("@LastName", txtPassword.Text)
                Dim userExists As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                If txtUsername.Text = "Austin" Then
                    MsgBox("Welcome to my household system")
                Else
                    If txtPassword.Text = "Ramuhashi" Then
                        MsgBox("welcome to my household system")
                    Else

                        MsgBox("I don't recogonize this loggins")

                        If userExists > 0 Then
                            MessageBox.Show("Login Successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    End If
                    DashBoard.TextBox1.Text = txtUsername.Text & " " & txtPassword.Text
                End If
                conn.Close()
                'MY_USER_VALIDATION.ShowDialog()
                'DashBoard.ShowDialog()

            End Using

            'End Using



            'Visible = False

            'Authentication.Show()

            'Try

            '    Dim username As String = txtFirstName.Text.Trim()

            '    Dim password As String = txtLastname.Text.Trim()

            '    ' Check if fields are empty

            '    If String.IsNullOrEmpty(username) OrElse String.IsNullOrEmpty(password) Then

            '        MessageBox.Show("Please fill in all fields.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            '        Exit Sub

            '    End If

            '    ' Check if username already exists

            '    If UsernameExists(username) Then

            '        MessageBox.Show("Username already exists. Please choose another one.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            '        Exit Sub

            '    End If



            '    Debug.WriteLine("entering button submit")

            '    Dim Regis As New regis() With {

            '    .UserName = txtFirstName.Text,

            '.Password = txtLastname.Text,

            '.EmailAddress = txtEmail.Text,

            '.Role = cmbxRole.SelectedItem.ToString(),

            '.FirstName = TextBox1.Text,

            '.LastName = TextBox2.Text}

            'Using conn As New OleDbConnection(connectionString)

            '        conn.Open()

            '        Dim tablename As String = "LoginDetails"

            '        Dim Cmd As New OleDbCommand($"INSERT INTO {tablename} ([UserName], [Password], [Email], [Role], [FirstName], [LastName]) VALUES (?, ?, ?, ?, ?, ?)", conn)


            '        Cmd.Parameters.AddWithValue("@Username", username)

            '        Cmd.Parameters.AddWithValue("@Password", password)

            '        Cmd.Parameters.AddWithValue("@Email", txtEmail.Text)

            '        Cmd.Parameters.AddWithValue("@Role", cmbxRole.SelectedItem.ToString())

            '        Cmd.Parameters.AddWithValue("@FirstName", TextBox1.Text)

            '        Cmd.Parameters.AddWithValue("@LastName", TextBox2.Text)


            '        Cmd.ExecuteNonQuery()

            '        MsgBox("You have Registered successfully!" & vbCrLf &

            '    "UserName: " & TxtFirstname.Text.ToString & vbCrLf &

            '    "Password: " & TxtLastName.Text.ToString & vbCrLf &

            '    "Email: " & TxtEmail.Text.ToString & vbCrLf &

            '    "FirstName: " & TextBox1.Text.ToString & vbCrLf &

            '    "LastName: " & TextBox2.Text.ToString & vbCrLf &

            '    "Role: " & cmbxRole.SelectedItem.ToString, vbInformation, "Register Confirmation")



            '    MessageBox.Show("Registration successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

            '        txtFirstName.Clear()

            '        txtLastname.Clear()

            '    End Using

            'Catch ex As OleDbException

            '    Debug.WriteLine($"Database error in btnSubmit_Click: {ex.Message}")

            '    Debug.WriteLine($"Stack Trace: {ex.StackTrace}")

            '    MessageBox.Show($"Error Saving registeration To database:  {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)


            'Catch ex As Exception

            '    Debug.WriteLine($"General error in Button Register: {ex.Message}")

            '    Debug.WriteLine($"Stack Trace: {ex.StackTrace}")

            '    MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            '    MessageBox.Show("Error saving control texts: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            'End Try

            ''LoadRegisterdatafromfile()

            'LoadRegisterdatafromfile()

            'conn.Close()

            'Debug.WriteLine("Exiting button Register")

            'ClearControls(Me)


        End Sub
    End Class



