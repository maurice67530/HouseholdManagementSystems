Imports System.IO
Imports System.Data.OleDb
Public Class Personnel
    Dim connec As New OleDbConnection(HouseHoldManagment_Module.connectionString)

    ' Create a ToolTip object
    Private toolTip As New ToolTip()
    Private toolTip1 As New ToolTip()


    Public Folderpath As String = "\\MUDAUMURANGI\Users\Murangi\Source\Repos\maurice67530\HouseholdManagementSystems\Personnel Pictures"



    Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles BtnSave.Click

        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            Try
                Dim selectedPath As String = OpenFileDialog1.FileName
                Dim imageName As String = Path.GetFileName(selectedPath)
                Dim destinationPath As String = Path.Combine(Folderpath, imageName)

                ' Save only the full UNC path to database for portability
                Dim dbFilePath As String = destinationPath

                Using conn As New OleDb.OleDbConnection(connectionString)
                    conn.Open()

                    ' Check if the image is already saved
                    Using checkCmd As New OleDb.OleDbCommand("SELECT COUNT(*) FROM Photos WHERE FilePath = ?", conn)
                        checkCmd.Parameters.AddWithValue("?", dbFilePath)
                        Dim count As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())

                        If count > 0 Then
                            MsgBox("This image has already been uploaded.", vbInformation, vbOKCancel)
                            Exit Sub
                        End If
                    End Using

                    ' Only copy if not already existing in destination folder
                    If Not Directory.Exists(Folderpath) Then
                        Directory.CreateDirectory(Folderpath)
                    End If

                    ' Optional: Check file existence in destination folder too
                    If Not File.Exists(destinationPath) Then
                        File.Copy(selectedPath, destinationPath, True)
                    End If

                    ' Save new record

                    Using cmd As New OleDb.OleDbCommand("INSERT INTO PersonalDetails ([Photo], [FirstName], [LastName], [DateOfBirth], [Gender], [Contact], [Email], [Role], [Age], [PostalCode], [MaritalStatus], [Dietary] ) VALUES ( ?, ?, ?, ?, ?, ?, ?, ? ,?, ?, ?, ?)", conn)
                        Dim person As New Person
                        cmd.Parameters.AddWithValue("?", dbFilePath)

                        cmd.Parameters.AddWithValue("?", TextBox1.Text)
                        cmd.Parameters.AddWithValue("?", TextBox2.Text)
                        cmd.Parameters.AddWithValue("?", DateTimePicker1.Value)
                        cmd.Parameters.AddWithValue("?", ComboBox3.SelectedItem.ToString())
                        cmd.Parameters.AddWithValue("?", TextBox3.Text)
                        cmd.Parameters.AddWithValue("?", TextBox4.Text)
                        cmd.Parameters.AddWithValue("?", ComboBox1.SelectedItem.ToString())
                        cmd.Parameters.AddWithValue("?", TextBox5.Text)
                        cmd.Parameters.AddWithValue("?", TextBox6.Text)
                        cmd.Parameters.AddWithValue("?", ComboBox2.SelectedItem.ToString())
                        cmd.Parameters.AddWithValue("?", ComboBox4.SelectedItem.ToString())


                        cmd.ExecuteNonQuery()
                        MsgBox(" You are now added as a member of the HoseHold Managment System!" & vbCrLf &
                              "FirstName: " & Person.FirstName & vbCrLf &
                              "LastName: " & Person.LastName & vbCrLf &
                              "Birth of Date:" & Person.DateOfBirth & vbCrLf &
                              "Gender: " & Person.Gender & vbCrLf &
                              "Contact Number: " & Person.Contact & vbCrLf &
                              "Email: " & Person.Email & vbCrLf &
                              "Role: " & Person.Role & vbCrLf &
                              "Age: " & Person.Age & vbCrLf &
                              "Postal Code: " & Person.postalcode & vbCrLf &
                              "Photo: " & Person.Photo & vbCrLf &
                              "Dietary: " & Person.Dietary & vbCrLf &
                              "Health Status: " & Person.MaritalStatus & vbCrLf, vbInformation, "Credentials  confirmation")
                    End Using
                End Using
                conn.Close()
                MessageBox.Show("Photo saved to database and network folder.")
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End If
        MessageBox.Show("Personnel information saved to Database successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)


        connec.Close()
        Debug.WriteLine("Exiting btnSubmit")
        'connec.Close()
        LoadData()
        ClearForm()
        ClearControls(Me)
    End Sub

    Private Sub Person_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadData()
        ' Initialize ToolTip properties (optional)
        toolTip.AutoPopDelay = 5000
        toolTip.InitialDelay = 500
        toolTip.ReshowDelay = 200
        toolTip.ShowAlways = True

        toolTip1.SetToolTip(BtnBack, "Back")
        'toolTip1.SetToolTip(BtnAddpicture, "Add a Picture")
        toolTip1.SetToolTip(BtnEdit, "Edit")
        toolTip1.SetToolTip(BtnDelete, "Delete")
        toolTip1.SetToolTip(BtnClear, "Clear")
        toolTip1.SetToolTip(BtnDailyTasks, "Daily tasks")
        toolTip1.SetToolTip(BtnSave, "Save")
        toolTip1.SetToolTip(Button1, "Refresh Table")

    End Sub
    ' Method to load data into the DataGridView
    Private Sub LoadData()
        Try
            connec.Open()
            Dim query As String = "SELECT * FROM PersonalDetails"
            Dim cmd As New OleDbCommand(query, connec)
            Dim adapter As New OleDbDataAdapter(cmd)
            Dim dt As New DataTable()
            adapter.Fill(dt)
            DataGridView1.DataSource = dt
        Catch ex As Exception
            MessageBox.Show("Failed to load data: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            connec.Close()
        End Try

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        ' When a row is clicked in the DataGridView

        ' Check if a valid row is clicked (not header)
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)

            ' Fill the form fields with the values from the selected row
            TextBox1.Text = row.Cells("FirstName").Value.ToString()
            TextBox2.Text = row.Cells("LastName").Value.ToString()
            DateTimePicker1.Value = Convert.ToDateTime(row.Cells("DateOfBirth").Value)
            TextBox4.Text = row.Cells("Email").Value.ToString()
            TextBox3.Text = row.Cells("Contact").Value.ToString()
            TextBox5.Text = row.Cells("Age").Value.ToString()
            ComboBox1.Text = row.Cells("Role").Value.ToString()
            ComboBox3.Text = row.Cells("Gender").Value.ToString()
            TextBox6.Text = row.Cells("PostalCode").Value.ToString()
            TextBox7.Text = row.Cells("Photo").Value.ToString()
            PictureBox1.ImageLocation = row.Cells("Photo").Value.ToString()
            ComboBox2.SelectedItem = row.Cells("MaritalStatus").Value.ToString()
            ComboBox4.SelectedItem = row.Cells("Dietary").Value.ToString()
        End If

    End Sub

    Private Sub BtnEdit_Click(sender As Object, e As EventArgs) Handles BtnEdit.Click
        Debug.WriteLine("Entering btnEdit")

        ' Ensure a row is selected in the DataGridView  
        If DataGridView1.SelectedRows.Count = 0 Then
            Debug.WriteLine("User canceled btnEdit")
            MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Return
        End If

        Try
            Debug.WriteLine("User confirmed btnEdit")

            Using connec As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                connec.Open()


                Dim FirstName As String = TextBox1.Text
                Dim LastName As String = TextBox2.Text
                Dim Gender As String = ComboBox3.SelectedItem.ToString
                Dim Contact As String = TextBox3.Text
                Dim Email As String = TextBox4.Text
                Dim Role As String = ComboBox1.SelectedItem.ToString
                Dim Age As String = TextBox5.Text
                Dim PostalCode As String = TextBox6.Text
                Dim MaritalStatus As String = ComboBox2.SelectedItem.ToString
                Dim DateOfBirth As String = DateTimePicker1.Value

                Dim Photo As String = TextBox7.Text
                Dim Dietary As String = ComboBox4.SelectedItem.ToString
                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Create an OleDbCommand to update the personnel data in the database  
                Dim cmd As New OleDbCommand("UPDATE [PersonalDetails] SET [FirstName] = ?, [LastName] = ?, [Gender] = ?, [Contact] = ?, [Email] = ?, [Role] = ?, [Age] = ?, [PostalCode] = ?, [MaritalStatus] = ?, [Photo] = ?, [Dietary] = ? WHERE ID = ?", connec)

                ' Set the parameter values from the UI controls  
                cmd.Parameters.AddWithValue("@FirstName", FirstName)
                cmd.Parameters.AddWithValue("@LastName", LastName)
                cmd.Parameters.AddWithValue("@Gender", Gender)
                cmd.Parameters.AddWithValue("@Contact", Contact)
                cmd.Parameters.AddWithValue("@Email", Email)
                cmd.Parameters.AddWithValue("@Role", Role)
                cmd.Parameters.AddWithValue("@Age", Age)
                cmd.Parameters.AddWithValue("@PostalCode", PostalCode)
                cmd.Parameters.AddWithValue("MaritalStatus", MaritalStatus)
                cmd.Parameters.AddWithValue("Photo", Photo)
                cmd.Parameters.AddWithValue("Dietary", Dietary)
                cmd.Parameters.AddWithValue(" DateOfBirth", DateOfBirth)
                cmd.Parameters.AddWithValue("ID", ID)

                ' Execute the SQL command to update the data
                cmd.ExecuteNonQuery()


                MsgBox("Personnel information updated!")


                LoadData()

                connec.Close()

            End Using
        Catch ex As OleDbException
            'Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            'MessageBox.Show($"Error updating personnel in database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'MessageBox.Show("Error saving personnel to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As FormatException
            Debug.WriteLine($"Format error in btnEdit:{ex.Message}")
            'Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            MessageBox.Show("Ensure all feilds are filled correctly.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch ex As Exception
            'MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine("Failed entering btnEdit_")
            'MessageBox.Show("Error saving personnel to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'Debug.WriteLine($"Stack Trace : {ex.StackTrace}")
            Debug.WriteLine($"An  Error has occured when Editing data from Database")
        End Try
        'connect.close
        connec.Close()
    End Sub

    Private Sub BtnDelete_Click(sender As Object, e As EventArgs) Handles BtnDelete.Click
        ' Check if a member is selected
        If String.IsNullOrWhiteSpace(TextBox8.Text) Then
            MessageBox.Show("Please select a member from the table before deleting.", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' Confirm before deletion
        Dim confirmDelete As DialogResult = MessageBox.Show("Are you sure you want to delete this member?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If confirmDelete = DialogResult.No Then
            Exit Sub
        End If

        Try
            connec.Open()
            Dim query As String = "DELETE FROM PersonalDetails WHERE ID = ?"

            Using cmd As New OleDbCommand(query, connec)
                cmd.Parameters.AddWithValue("@ID", CInt(TextBox8.Text))
                cmd.ExecuteNonQuery()
            End Using

            MessageBox.Show("Member deleted successfully.")
            LoadData()
            ClearForm()

        Catch ex As Exception
            MessageBox.Show("Delete failed: " & ex.Message)
        Finally
            connec.Close()
        End Try
    End Sub

    Private Sub BtnBack_Click(sender As Object, e As EventArgs) Handles BtnBack.Click
        Me.Visible = False
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim row As DataGridViewRow = DataGridView1.SelectedRows(0)
            If Not row.IsNewRow Then
                TextBox8.Text = row.Cells("ID").Value.ToString()
                TextBox1.Text = row.Cells("FirstName").Value.ToString()
                TextBox2.Text = row.Cells("LastName").Value.ToString()
                TextBox3.Text = row.Cells("Contact").Value.ToString()
                TextBox4.Text = row.Cells("Email").Value.ToString()
                TextBox5.Text = row.Cells("Age").Value.ToString()
                ComboBox2.SelectedItem = row.Cells("MaritalStatus").Value.ToString()
                ComboBox1.SelectedItem = row.Cells("Role").Value.ToString()
                ComboBox3.SelectedItem = row.Cells("Gender").Value.ToString()
                TextBox6.Text = row.Cells("PostalCode").Value.ToString()
                TextBox7.Text = row.Cells("Photo").Value.ToString()
                PictureBox1.ImageLocation = row.Cells("Photo").Value.ToString()
                DateTimePicker1.Value = row.Cells("DateOfBirth").Value.ToString()
                ComboBox4.SelectedItem = row.Cells("Dietary").Value.ToString()

            End If
        End If
    End Sub
    Private Sub ClearForm()
        TextBox8.Clear()
        TextBox1.Clear()
        TextBox2.Clear()
        'DateTimePicker1.CLEAR
        TextBox4.Clear()
        TextBox3.Clear()
        TextBox5.Clear()
        'ComboBox1.CLEAR
        'ComboBox3.CLEAR
        TextBox6.Clear()
        'ComboBox2.CLEAR

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox8.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
    End Sub
    Public Sub ClearControls(ByVal FORM As Form)
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
    'Dim OpenFileDialog As New OpenFileDialog()
    'OpenFileDialog.Filter = "Bitmaps (*.jpg)|*.jpg"
    'If OpenFileDialog.ShowDialog() = DialogResult.OK Then
    '    PictureBox1.ImageLocation = OpenFileDialog.FileName
    '    TextBox7.Text = OpenFileDialog.FileName
    'End If




    Private Sub BtnDailyTasks_Click(sender As Object, e As EventArgs) Handles BtnDailyTasks.Click
        chores.ShowDialog()
    End Sub

    Private Sub BtnClear_Click(sender As Object, e As EventArgs) Handles BtnClear.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox8.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        TextBox7.Text = ""
        ComboBox4.Text = ""
        PictureBox1.ImageLocation = ""
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        LoadData()
    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub




End Class