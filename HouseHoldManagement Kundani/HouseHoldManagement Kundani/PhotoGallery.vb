Imports System.IO
Imports System.Data.OleDb
Public Class PhotoGallery
    Private SelectedImagePath As String = " "

    Public Folderpath As String = "\\MUDAUMURANGI\Users\Murangi\Source\Repos\maurice67530\HouseholdManagementSystems\Photo Gallery"

    Dim Photo As New Photo_Gallery

    Private Sub PhotoGallery_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckDatabaseConnection(statusLabel)


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


            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
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
        Using conn As New OleDb.OleDbConnection(HouseHoldManagment_Module.connectionString)
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
            Debug.WriteLine("Populate Datagridview: DataGridView populated successfully.")
            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                conn.Open()

                Dim tableName As String = "Photos"
                Dim query As String = $"SELECT * FROM [{tableName}]"  ' [] are useful in case of spaces in table names

                Using cmd As New OleDbCommand(query, conn)
                    Using da As New OleDbDataAdapter(cmd)
                        Dim dt As New DataTable
                        da.Fill(dt)
                        DataGridView1.DataSource = dt
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine("Failed to populate DataGridView")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error loading photos data from database: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
                    Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
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

    'Dim folderPath As String = Application.StartupPath & "\Photo Gallery\"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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
                    Using cmd As New OleDb.OleDbCommand("INSERT INTO Photos ([Description], [FilePath], [DateAdded], [FamilyMember], [Photographer], [Album]) VALUES (?, ?, ?, ?, ?, ?)", conn)
                        cmd.Parameters.AddWithValue("?", TextBox2.Text)
                        cmd.Parameters.AddWithValue("?", dbFilePath)
                        cmd.Parameters.AddWithValue("?", DateTimePicker1.Value)
                        cmd.Parameters.AddWithValue("?", ComboBox1.Text)
                        cmd.Parameters.AddWithValue("?", TextBox3.Text)
                        cmd.Parameters.AddWithValue("?", ComboBox2.Text)
                        cmd.ExecuteNonQuery()
                    End Using
                End Using

                MessageBox.Show("Photo saved to database and network folder.")
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End If
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

        HouseHoldManagment_Module.FilterPhoto(selectedFamilyMember) ', selectedDateAdded)
    End Sub
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Try
            If DataGridView1.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

                ' Populate controls
                TextBox2.Text = selectedRow.Cells("Description").Value?.ToString()
                DateTimePicker1.Text = selectedRow.Cells("DateAdded").Value?.ToString()
                ComboBox1.Text = selectedRow.Cells("FamilyMember").Value?.ToString()
                TextBox3.Text = selectedRow.Cells("Photographer").Value?.ToString()
                ComboBox2.Text = selectedRow.Cells("Album").Value?.ToString()

                ' Try to load the image from the UNC path in FilePath
                Dim filePath As String = selectedRow.Cells("FilePath").Value?.ToString()
                If Not String.IsNullOrWhiteSpace(filePath) AndAlso File.Exists(filePath) Then
                    PictureBox1.ImageLocation = filePath
                    PictureBox1.Image = Image.FromFile(filePath)
                Else
                    PictureBox1.Image = Nothing
                    MessageBox.Show("Image not found at " & filePath)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Error selecting data: " & ex.Message)
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

    Private loadedImagePath As String = String.Empty ' To store selected image path
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        'Dim openFileDialog As New OpenFileDialog()
        'openFileDialog.Title = "Select an image"
        'openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp"

        'If openFileDialog.ShowDialog() = DialogResult.OK Then
        '    loadedImagePath = openFileDialog.FileName
        '    PictureBox1.Image = Image.FromFile(loadedImagePath)
        'End If


        ' Load last photo from database
        Try
            conn.Open()
            Dim cmd As New OleDbCommand("SELECT TOP 1 * FROM Photos ORDER BY ID DESC", conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            If reader.Read() Then
                Dim imageName As String = reader("FilePath").ToString()
                Dim fullPath As String = Path.Combine(folderPath, imageName)

                If File.Exists(fullPath) Then
                    If PictureBox1.Image IsNot Nothing Then PictureBox1.Image.Dispose()
                    PictureBox1.Image = Image.FromFile(fullPath)
                Else
                    MessageBox.Show("Image not found in folder.")
                End If
            Else
                MessageBox.Show("No photo found.")
            End If

            reader.Close()
            conn.Close()
        Catch ex As Exception
            conn.Close()
            MessageBox.Show("Error loading image: " & ex.Message)
        End Try

    End Sub
    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click

    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        DataGridView1.Sort(DataGridView1.Columns("DateAdded"), System.ComponentModel.ListSortDirection.Ascending)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs)
        Try
            Using conn As New OleDbConnection(connectionString)
                conn.Open()
                Dim selectCmd As New OleDbCommand("SELECT [ID], [FilePath] FROM [Photos]", conn)
                Dim da As New OleDbDataAdapter(selectCmd)
                Dim dt As New DataTable()
                da.Fill(dt)

                For Each row As DataRow In dt.Rows
                    Dim fileName As String = row("FilePath").ToString()
                    If Not fileName.StartsWith("\\MUDAUMURANGI") Then
                        Dim newPath = System.IO.Path.Combine(Folderpath, fileName)
                        Dim updateCmd As New OleDbCommand("UPDATE [Photos] SET [FilePath]=? WHERE [ID]=?", conn)
                        updateCmd.Parameters.AddWithValue("?", newPath)
                        updateCmd.Parameters.AddWithValue("?", row("ID"))
                        updateCmd.ExecuteNonQuery()
                    End If
                Next
            End Using
            MessageBox.Show("FilePath values updated!", "Maintenance", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error updating paths: " & ex.Message)
        End Try
    End Sub

    Private Sub refresh_Click(sender As Object, e As EventArgs) Handles refresh.Click
        LoadPhotodataFromDatabase()

        TextBox2.Text = " "

        'DateTimePicker1.Text = " "

        ComboBox1.Text = " "

        TextBox3.Text = " "

        ComboBox2.Text = " "

        Label10.Text = " "


    End Sub
End Class