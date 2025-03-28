Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Drawing
Imports System.IO

Public Class PHOTO_GALLERY
    Private _connection As OleDbConnection

    Private Sub InitializeConnection()
        _connection = New OleDbConnection(Module1.connectionString)
    End Sub

    Private Sub PHOTO_GALLERY_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            InitializeConnection()
            ToolTip1.SetToolTip(Button1, "Dashboard")
            LoadPhotoGalleryFromDatabase()
        Catch ex As Exception
            MessageBox.Show($"Error initializing form: {ex.Message}", "Initialization Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub LoadPhotoGalleryFromDatabase()
        Try
            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()
                Dim cmd As New OleDbCommand("SELECT * FROM PhotoAlbum", conn)
                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)
                DataGridView1.DataSource = dt
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error loading photo gallery data: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function ValidatePhotoData() As Boolean
        If String.IsNullOrEmpty(DateTimePicker1.Text) Then
            MessageBox.Show("Please select a date", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If
        If String.IsNullOrEmpty(ComboBox1.Text) Then
            MessageBox.Show("Please select a payment method", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If
        If String.IsNullOrEmpty(TextBox2.Text) OrElse Not IsNumeric(TextBox2.Text) Then
            MessageBox.Show("Please enter a valid price", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If
        If String.IsNullOrEmpty(TextBox3.Text) OrElse Not IsNumeric(TextBox3.Text) Then
            MessageBox.Show("Please enter a valid number of photos", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If
        If String.IsNullOrEmpty(TextBox4.Text) OrElse Not IsNumeric(TextBox4.Text) Then
            MessageBox.Show("Please enter a valid photo ID", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If
        If String.IsNullOrEmpty(TextBox5.Text) Then
            MessageBox.Show("Please enter a title", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If
        If String.IsNullOrEmpty(ListBox1.Text) Then
            MessageBox.Show("Please select or add a photo", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If
        Return True
    End Function

    Private Sub SavePhotoData()
        If Not ValidatePhotoData() Then Return

        Try
            Dim photo As New photogallery With {
                .DateoftheDay = DateTimePicker1.Text,
                .paymentMethod = ComboBox1.Text,
                .priceperitem = TextBox2.Text,
                .NumberOfPhototaken = TextBox3.Text,
                .photoID = TextBox4.Text,
                .Title = TextBox5.Text,
                .Description = TextBox6.Text,
                .filepath = ListBox1.Text,
                .Category = ComboBox4.Text,
                .Tag = TextBox8.Text,
                .Picture = ListBox1.Text
            }

            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()
                Dim cmd As New OleDbCommand("INSERT INTO [photoAlbum] ([DateOftheDay],[paymentMethod],[PriceperItem],[NumberOfPhototaken],[photoID],[Title],[Description],[filepath],[Tag],[Category],[Picture]) VALUES (@DateOftheDay,@paymentMethod,@priceperitem,@NumberOfPhototaken,@photoID,@Title,@Description,@filepath,@Tag,@Category,@Picture)", conn)

                cmd.Parameters.AddWithValue("@DateOftheday", photo.DateoftheDay)
                cmd.Parameters.AddWithValue("@paymentMethod", photo.paymentMethod)
                cmd.Parameters.AddWithValue("@priceperitem", photo.priceperitem)
                cmd.Parameters.AddWithValue("@NumberOfPhototaken", photo.NumberOfPhototaken)
                cmd.Parameters.AddWithValue("@PhotoID", photo.photoID)
                cmd.Parameters.AddWithValue("@Title", photo.Title)
                cmd.Parameters.AddWithValue("@Description", photo.Description)
                cmd.Parameters.AddWithValue("@Filepath", photo.filepath)
                cmd.Parameters.AddWithValue("@Category", photo.Category)
                cmd.Parameters.AddWithValue("@Tag", photo.Tag)
                cmd.Parameters.AddWithValue("@Picture", photo.Picture)

                cmd.ExecuteNonQuery()
                MessageBox.Show("Photo gallery entry saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                LoadPhotoGalleryFromDatabase()
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error saving photo data: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SaveToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem1.Click
        SavePhotoData()
    End Sub

    'Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    '    SavePhotoData()
    'End Sub

    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        If DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a record to delete", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim result = MessageBox.Show("Are you sure you want to delete this photo entry?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Try
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim photoId As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value)

                Using conn As New OleDbConnection(Module1.connectionString)
                    conn.Open()
                    Dim cmd As New OleDbCommand("DELETE FROM [PhotoAlbum] WHERE [ID] = @ID", conn)
                    cmd.Parameters.AddWithValue("@ID", photoId)
                    cmd.ExecuteNonQuery()
                    MessageBox.Show("Photo entry deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    LoadPhotoGalleryFromDatabase()
                End Using
            Catch ex As Exception
                MessageBox.Show($"Error deleting photo entry: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub EditToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditToolStripMenuItem.Click
        If DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a record to edit", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
            Dim photoId As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value)

            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()
                Dim cmd As New OleDbCommand("UPDATE PhotoAlbum SET [DateOftheday]=@DateOftheday, [paymentMethod]=@paymentMethod, [Priceperitem]=@Priceperitem, [NumberOfPhototaken]=@NumberOfPhototaken, [PhotoID]=@PhotoID, [Title]=@Title, [Description]=@Description, [Filepath]=@Filepath, [Tag]=@Tag, [Category]=@Category, [Picture]=@Picture WHERE [ID]=@ID", conn)

                cmd.Parameters.AddWithValue("@DateOftheday", DateTimePicker1.Text)
                cmd.Parameters.AddWithValue("@paymentMethod", ComboBox1.Text)
                cmd.Parameters.AddWithValue("@Priceperitem", TextBox2.Text)
                cmd.Parameters.AddWithValue("@NumberOfPhototaken", TextBox3.Text)
                cmd.Parameters.AddWithValue("@PhotoID", TextBox4.Text)
                cmd.Parameters.AddWithValue("@Title", TextBox5.Text)
                cmd.Parameters.AddWithValue("@Description", TextBox6.Text)
                cmd.Parameters.AddWithValue("@Filepath", ListBox1.Text)
                cmd.Parameters.AddWithValue("@Tag", TextBox8.Text)
                cmd.Parameters.AddWithValue("@Category", ComboBox4.Text)
                cmd.Parameters.AddWithValue("@Picture", ListBox1.Text)
                cmd.Parameters.AddWithValue("@ID", photoId)

                cmd.ExecuteNonQuery()
                MessageBox.Show("Photo entry updated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                LoadPhotoGalleryFromDatabase()
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error updating photo entry: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SavePhotoFromFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SavePhotoFromFileToolStripMenuItem.Click
        OpenFileDialog1.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif"
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Try
                PictureBox1.Image = Image.FromFile(OpenFileDialog1.FileName)
                ListBox1.Text = OpenFileDialog1.FileName
            Catch ex As Exception
                MessageBox.Show($"Error loading image: {ex.Message}", "Image Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Try
            'imagePaths.Clear()
            'PictureBox1.Image = Nothing
            'currentImageIndex = 0
            'Label12.Text = String.Empty

            'If ComboBox4.SelectedIndex <> -1 Then
            '    Dim selectedDescription As String = ComboBox4.SelectedItem.ToString()
            '    LoadPhotosFromDatabase(selectedDescription)

            '    If imagePaths.Count > 0 Then
            ''        Timer1.Start()
            'End If
            'End If
        Catch ex As Exception
            MessageBox.Show($"Error loading photos for category: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub LoadPhotosFromDatabase(description As String)
        Try
            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()
                Dim query As String = "SELECT FilePath, DateOftheday FROM PhotoAlbum WHERE Category = @Category"
                Using cmd As New OleDbCommand(query, conn)
                    cmd.Parameters.AddWithValue("@Category", description)
                    Using reader As OleDbDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim photoPath As String = reader("FilePath").ToString()
                            If File.Exists(photoPath) Then
                                'imagePaths.Add(photoPath)
                            End If
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error loading photos: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
    '    'Try
    '    If imagePaths.Count > 0 AndAlso currentImageIndex < imagePaths.Count Then
    '        Dim photoPath As String = imagePaths(currentImageIndex)
    '        If File.Exists(photoPath) Then
    '            PictureBox1.Image = Image.FromFile(photoPath)
    '            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
    '            Label12.Text = $"Date Taken: {DateTimePicker1.Value:MM/dd/yyyy}"
    '        End If

    '    '        currentImageIndex = (currentImageIndex + 1) Mod imagePaths.Count
    '    '    End If
    '    Catch ex As Exception
    '        MessageBox.Show($"Error displaying photo: {ex.Message}", "Display Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        DashBoard.Show()
        Me.Close()
    End Sub
End Class