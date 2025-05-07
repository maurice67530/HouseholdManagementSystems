Imports System.IO
Imports System.Data.OleDb
Public Class PhotoGallery

    Private Sub PhotoGallery_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        PopulateComboboxFromDatabase(ComboBox1)
        LoadPhotodataFromDatabase()

        ToolTip1.SetToolTip(Button1, "Save")
        ToolTip1.SetToolTip(Button2, "Update")
        ToolTip1.SetToolTip(Button3, "Filter")
        ToolTip1.SetToolTip(Button4, "Delete")
        ToolTip1.SetToolTip(Button5, "Search by Family Member")
        ToolTip1.SetToolTip(Button6, "Stop Images")
        ToolTip1.SetToolTip(Button8, "Upload Image")
        ToolTip1.SetToolTip(Button7, "Sort")

    End Sub
    Public Sub eish()
        Dim items As New List(Of String) From {
               "Church",
               "Soccer",
               "Christmas",
               "Family"
           }

        ' Set the ComboBox data source
        ComboBox2.DataSource = items
    End Sub
    Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Debug.WriteLine("Entering button update click")
        If DataGridView1.SelectedRows.Count = 0 Then
            Debug.WriteLine("User confirmed update")
            MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Debug.WriteLine($"Format error in button update:")
            Debug.WriteLine("Updating data: data updated")
            Dim Description As String = TextBox2.Text
            Dim FilePath As String = PictureBox1.ImageLocation
            Dim DateAdded As String = DateTimePicker1.Value
            Dim FamilyMember As String = ComboBox1.SelectedItem.ToString()
            Dim Photographer As String = TextBox3.Text
            Dim Album As String = ComboBox2.SelectedItem.ToString()


            Using conn As New OleDbConnection(Rinae.connectionString)
                conn.Open()

                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Create an OleDbCommand to update the personnel data in the database  
                Dim cmd As New OleDbCommand("UPDATE [Photos] SET [Description] = ?, [FilePath] = ?, [DateAdded] = ?, [FamilyMember] = ?, [Photographer] = ?, [Album] = ? WHERE [ID] = ?", conn)

                ' Set the parameter values from the UI controls  

                'cmd.Parameters.AddWithValue("@PhotoID", PhotoID.textbox1.Text)
                cmd.Parameters.AddWithValue("@Description", Description)
                cmd.Parameters.AddWithValue("@FilePath", FilePath)
                cmd.Parameters.AddWithValue("@DateAdded", DateAdded)
                cmd.Parameters.AddWithValue("@FamilyMember", FamilyMember)
                cmd.Parameters.AddWithValue("@Photographer", Photographer)
                cmd.Parameters.AddWithValue("@Album", Album)
                cmd.Parameters.AddWithValue("@ID", ID)

                cmd.ExecuteNonQuery()

                MsgBox("Photo Updated Successfuly!", vbInformation, "Update Confirmation")
                LoadPhotodataFromDatabase()
                ' HouseHold.ClearControls(Me)
            End Using
        Catch ex As OleDbException
            Debug.WriteLine("User cancelled update")
            Debug.WriteLine("Unexpected error in button update")
            Debug.Write($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error updating Photos in database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Debug.WriteLine("Exiting button update")

    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim SearchTerm As String = TextBox4.Text
        'Dim connString As String()
        Using conn As New OleDb.OleDbConnection(Rinae.connectionString)
            conn.Open()

            Dim cmd As New OleDb.OleDbCommand("SELECT * FROM Photos WHERE Familymember LIKE ?", conn)
            cmd.Parameters.AddWithValue("Familymember", "%" & SearchTerm & "%")
            Dim dataAdapter As New OleDbDataAdapter(cmd)
            Dim dt As New DataTable()
            dataAdapter.Fill(dt)
            DataGridView1.DataSource = dt
        End Using
    End Sub
    Public Sub LoadPhotodataFromDatabase()
        Try
            '  Dim dataTable As DataTable = HouseHold.GetData("SELECT * FROM Expense")
            ' DataGridView1.DataSource = DataTable
            Debug.WriteLine("Populate Datagridview: Datagridview populated successfully.")
            Using conn As New OleDbConnection(Rinae.connectionString)
                conn.Open()

                Dim tableName As String = "Photos"

                Dim cmd As New OleDbCommand($"SELECT*FROM {tableName}", conn)

                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)
                DataGridView1.DataSource = dt
            End Using
        Catch ex As Exception
            Debug.WriteLine("Failed to populate Gatagridview")
            Debug.Write($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error Loading photos data from database: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
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
            Dim confirmationResult As DialogResult = MessageBox.Show("Are you sure you want to delete this Photos?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If confirmationResult = DialogResult.Yes Then
                Debug.WriteLine("User confirmation deletion.")
                ' Proceed with deletion  
                Try
                    Debug.WriteLine("Format errors in button delete")
                    Debug.WriteLine("Deleting data: Data delected")
                    Debug.WriteLine("Stack Trace: {ex.StackTrace}")
                    Using conn As New OleDbConnection(Rinae.connectionString)
                        conn.Open()

                        ' Create the delete command  
                        Dim cmd As New OleDbCommand("DELETE FROM [Photos] WHERE [ID] = ?", conn)
                        cmd.Parameters.AddWithValue("@ID", PhotoID) ' Primary key for matching record  

                        ' Execute the delete command  
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show("Photos deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ' Optionally refresh DataGridView or reload from database  
                            ' PopulateDataGridView()
                        Else
                            MessageBox.Show("No Photo was deleted. Please check if the ID exists.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End Using
                    LoadPhotodataFromDatabase()
                Catch ex As Exception
                    Debug.WriteLine("Failed to delete data")
                    Debug.Write($"Stack Trace: {ex.StackTrace}")
                    MessageBox.Show($"An error occurred while deleting the Photos: {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            Debug.WriteLine("User cancelled deletion")
            MessageBox.Show("Please select Photos to delete.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

        Debug.WriteLine("Exiting button delete")

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            ''validate due date is not in the past
            'If TaskForm.DateTimePicker1.Value = DateTime.Now Then
            '    MessageBox.Show("Due Date cannot be in the past", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            '    Return
            'End If

            ''validate priority is selected
            'If String.IsNullOrEmpty(TextBox1.Text) Then
            '    MessageBox.Show("Please Enter a Title for the Photos", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            '    Return
            'End If

            Using conn As New OleDbConnection(Rinae.connectionString)

                conn.Open()
                Dim cmd As New OleDbCommand($"INSERT INTO Photos ([Description], [FilePath], [DateAdded], [FamilyMember], [Photographer], [Album]) VALUES (?, ?, ?, ?, ?, ?)", conn)

                cmd.Parameters.Clear()
                'cmd.Parameters.AddWithValue("@PhotoID", PhotoID.textbox1.Text)
                cmd.Parameters.AddWithValue("@Description", TextBox2.Text)
                cmd.Parameters.AddWithValue("@FilePath", PictureBox1.ImageLocation)
                cmd.Parameters.AddWithValue("@DateAdded", DateTimePicker1.Value)
                cmd.Parameters.AddWithValue("@FamilyMember", ComboBox1.SelectedItem)
                cmd.Parameters.AddWithValue("@Photographer", TextBox3.Text)
                cmd.Parameters.AddWithValue("@Album", ComboBox2.SelectedItem)

                cmd.ExecuteNonQuery()
                conn.Close()
                LoadPhotodataFromDatabase()
                MsgBox("Photo Added!" & vbCrLf &
                       "Description: " & TextBox2.Text & vbCrLf &
                        "FilePath: " & PictureBox1.ImageLocation & vbCrLf &
                        "DateAdded: " & DateTimePicker1.Value & vbCrLf &
                        "FamilyMember: " & ComboBox1.SelectedItem & vbCrLf &
                        "Album: " & ComboBox2.SelectedItem & vbCrLf &
                        "Photographer: " & TextBox3.Text)

            End Using
        Catch ex As OleDbException
            Debug.WriteLine($"Database error: {ex.Message}")
            Debug.Write($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("Error saving Photos to database: Please check the connectivity." & ex.Message & vbNewLine & ex.StackTrace, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Debug.WriteLine($"General error: {ex.Message}")

            MessageBox.Show("Unexpected Error: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
        End Try
        'Tags as Photo Day in Calender
        'Ndamu.MarkPhotoDay(DateTimePicker1.Text, TextBox2.Text)
    End Sub
    Public Sub PopulateComboboxFromDatabase(ByRef comboBox As ComboBox)
        Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
        Try
            Debug.WriteLine("Populating combobox: combobox populated from database")
            'open the database connection
            conn.Open()

            'retrieve the firstname and surname columns from the personaldetails tabel
            Dim query As String = "SELECT FirstName, LastName FROM PersonalDetails"
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            'bind the retrieved data to the combobox
            ComboBox1.Items.Clear()
            While reader.Read()
                ComboBox1.Items.Add($"{reader("FirstName")} {reader("LastName")}")
            End While

            'close the database
            reader.Close()

        Catch ex As Exception
            Debug.WriteLine("Failed to initialize combobox")
            Debug.Write($"Stack Trace: {ex.StackTrace}")
            'handle any exeptions that may occur
            MessageBox.Show($"Error: {ex.Message}")
        Finally
            'close the database connection
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
        Debug.WriteLine("Done with populating combobox from database")
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        '  Dim selectedDateAdded As String = If(DateTimePicker1.Text IsNot Nothing, DateTimePicker1.Text.ToString(), "")
        Dim selectedFamilyMember As String = If(ComboBox1.SelectedItem IsNot Nothing, ComboBox1.SelectedItem.ToString(), "")

        Cruwza.FilterPhoto(selectedFamilyMember) ', selectedDateAdded)
    End Sub
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Try
            Debug.WriteLine("Selecting data in the GDV: Data selected")
            If DataGridView1.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

                ' Load the data from the selected row into UI controls  
                TextBox2.Text = selectedRow.Cells("Description").Value.ToString()
                PictureBox1.ImageLocation = selectedRow.Cells("FilePath").Value.ToString()
                DateTimePicker1.Text = selectedRow.Cells("DateAdded").Value.ToString()
                ComboBox1.Text = selectedRow.Cells("FamilyMember").Value.ToString()
                TextBox3.Text = selectedRow.Cells("Photographer").Value.ToString()

            End If

            If DataGridView1.SelectedRows.Count > 0 Then
                Dim FilePath As String = DataGridView1.SelectedRows(0).Cells("FilePath").Value.ToString()

                If System.IO.File.Exists(FilePath) Then
                    PictureBox1.Image = Image.FromFile(FilePath)
                Else
                    PictureBox1.Image = Nothing
                    MessageBox.Show("IMAGE NOT FOUND")
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine("Data not selected: Error")
            Debug.Write($"Stack Trace: {ex.StackTrace}")
        End Try
    End Sub


    'Dim currentImageIndex As Integer = 0
    'Dim imagePaths As List(Of String) = New List(Of String)()

    'Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
    '    Timer1.Start()
    '    Timer1.Interval = 2000

    '    ' Ensure there are images to display
    '    If imagePaths.Count > 0 Then

    '        ' Check if the current image index is within the list range
    '        If currentImageIndex < imagePaths.Count Then

    '            ' Display the current image           
    '            Dim photoPath As String = imagePaths(currentImageIndex)

    '            ' Check if the photo exists
    '            If File.Exists(photoPath) Then
    '                PictureBox1.Image = Image.FromFile(photoPath)
    '                PictureBox1.SizeMode = PictureBoxSizeMode.Zoom

    '            Else
    '                MessageBox.Show("Image not found: " & photoPath)
    '            End If
    '            ' Optionally, display the date taken (this assumes you want to display the date of the current image)
    '            ' This will need to be adjusted to fetch the DateTaken from the database for each image
    '            ' Currently, we're assuming you fetch the photo's DateTaken in the LoadPhotosFromDatabase function.

    '            Dim dateTaken As String = DateTimePicker1.Value
    '            Label10.Text = "Date Taken: " & dateTaken
    '        End If
    '        ' Move to the next image in the list
    '        currentImageIndex += 1

    '        ' Reset to the first image once all images have been shown
    '        If currentImageIndex >= imagePaths.Count Then
    '            currentImageIndex = 0
    '        End If
    '    End If
    'End Sub
    '' Function to fetch photo paths from the database based on the selected description
    'Private Sub LoadPhotosFromDatabase(description As String)
    '    Try
    '        Using conn As New OleDbConnection(Cruwza.connectionString)

    '            Dim query As String = "SELECT FilePath, DateAdded FROM Photos WHERE Album = @Album"

    '            Using cmd As New OleDbCommand(query, conn)
    '                cmd.Parameters.AddWithValue("@Album", description)

    '                ' Open the database connection
    '                conn.Open()
    '                ' Execute the query and retrieve data
    '                Dim reader As OleDbDataReader = cmd.ExecuteReader()
    '                ' Loop through the results and populate the imagePaths list
    '                While reader.Read()

    '                    Dim photoPath As String = reader("FilePath").ToString()
    '                    imagePaths.Add(photoPath)
    '                End While
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '        MessageBox.Show("Error: " & ex.Message)
    '    Finally
    '        ' Close the database connection
    '        'conn.Close()
    '    End Try
    'End Sub
    'Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
    '    Timer1.Stop()
    'End Sub
    'Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs)
    '    ' Clear previous images and reset variables
    '    imagePaths.Clear()

    '    PictureBox1.Image = Nothing
    '    currentImageIndex = 0
    '    Label10.Text = String.Empty

    '    ' Check if a description is selected

    '    If ComboBox2.SelectedIndex <> -1 Then

    '        ' Get the selected description from ComboBox
    '        Dim selectedDescription As String = ComboBox2.SelectedItem.ToString()

    '        ' Get the photo paths based on the selected description
    '        LoadPhotosFromDatabase(selectedDescription)

    '        ' Start the timer to slide through the photos
    '        If imagePaths.Count > 0 Then
    '            Timer1.Start()
    '        End If
    '    End If
    '    'LoadFilteredMealPlan()
    'End Sub














    ' Timer variables
    Dim currentImageIndex As Integer = 0
    Dim imagePaths As List(Of String) = New List(Of String)()
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Timer1.Start()
        Timer1.Interval = 2000

        ' Ensure there are images to display
        If imagePaths.Count > 0 Then

            ' Check if the current image index is within the list range
            If currentImageIndex < imagePaths.Count Then

                ' Display the current image           
                Dim photoPath As String = imagePaths(currentImageIndex)

                ' Check if the photo exists
                If File.Exists(photoPath) Then
                    PictureBox1.Image = Image.FromFile(photoPath)
                    PictureBox1.SizeMode = PictureBoxSizeMode.Zoom

                Else
                    MessageBox.Show("Image not found: " & photoPath)
                End If
                ' Optionally, display the date taken (this assumes you want to display the date of the current image)
                ' This will need to be adjusted to fetch the DateTaken from the database for each image
                ' Currently, we're assuming you fetch the photo's DateTaken in the LoadPhotosFromDatabase function.

                Dim dateTaken As String = DateTimePicker1.Value
                Label10.Text = "Date Taken: " & dateTaken
            End If
            ' Move to the next image in the list
            currentImageIndex += 1

            ' Reset to the first image once all images have been shown
            If currentImageIndex >= imagePaths.Count Then
                currentImageIndex = 0
            End If
        End If
    End Sub
    ' Function to fetch photo paths from the database based on the selected description
    Private Sub LoadPhotosFromDatabase(description As String)
        Try
            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)

                Dim query As String = "SELECT FilePath, DateAdded FROM Photos WHERE Album = @Album"

                Using cmd As New OleDbCommand(query, conn)
                    cmd.Parameters.AddWithValue("@Album", description)

                    ' Open the database connection
                    conn.Open()
                    ' Execute the query and retrieve data
                    Dim reader As OleDbDataReader = cmd.ExecuteReader()
                    ' Loop through the results and populate the imagePaths list
                    While reader.Read()

                        Dim photoPath As String = reader("FilePath").ToString()
                        imagePaths.Add(photoPath)
                    End While
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            ' Close the database connection
            'conn.Close()
        End Try
    End Sub
    ' Stop the timer when you no longer want to slide the images
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Timer1.Stop()
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        ' Clear previous images and reset variables
        imagePaths.Clear()

        PictureBox1.Image = Nothing
        currentImageIndex = 0
        Label10.Text = String.Empty

        ' Check if a description is selected

        If ComboBox2.SelectedIndex <> -1 Then

            ' Get the selected description from ComboBox
            Dim selectedDescription As String = ComboBox2.SelectedItem.ToString()

            ' Get the photo paths based on the selected description
            LoadPhotosFromDatabase(selectedDescription)

            ' Start the timer to slide through the photos
            If imagePaths.Count > 0 Then
                Timer1.Start()
            End If
        End If
        'LoadFilteredMealPlan()

    End Sub












    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim OpenFileDialog As New OpenFileDialog()
        OpenFileDialog.Filter = "Bitmaps (*.jpg)|*.jpg"
        If OpenFileDialog.ShowDialog() = DialogResult.OK Then
            PictureBox1.ImageLocation = OpenFileDialog.FileName
            TextBox5.Text = OpenFileDialog.FileName
        End If
    End Sub
    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click

    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        DataGridView1.Sort(DataGridView1.Columns("DateAdded"), System.ComponentModel.ListSortDirection.Ascending)
    End Sub
End Class