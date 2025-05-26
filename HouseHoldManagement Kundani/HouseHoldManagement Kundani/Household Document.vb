Imports System.Data.OleDb
Imports System.IO

Public Class Household_Document


    Private SelectedImagePath As String = " "

    Public Folderpath As String = "\\MUDAUMURANGI\Users\Murangi\Source\Repos\maurice67530\HouseholdManagementSystems\House Hold Documents"
    'Private uploadPath As String = Path.Combine(Application.StartupPath, "Uploads")
    Public Property conn As New OleDbConnection(connectionString)

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        If TextBox1.Text = "" OrElse TextBox2.Text = "" Then
            MsgBox("Fill in the missing fields", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Dim ofd As New OpenFileDialog()
        ofd.Filter = "Documents|*.pdf;*.docx;*.xlsx;*.jpg;*.png|All files|*.*"

        If ofd.ShowDialog() = DialogResult.OK Then
            Dim sourcePath As String = ofd.FileName
            Dim fileName As String = IO.Path.GetFileName(sourcePath)

            ' Define your network folder and category subfolder
            Dim networkFolder As String = "\\MUDAUMURANGI\Users\Murangi\Source\Repos\maurice67530\HouseholdManagementSystems\House Hold Documents" ' <-- Replace with your actual path
            Dim categoryFolder As String = Path.Combine(networkFolder, ComboBox1.Text)

            ' Ensure the category folder exists
            Directory.CreateDirectory(categoryFolder)

            ' Build destination path and copy file
            Dim destinationPath As String = Path.Combine(categoryFolder, fileName)
            File.Copy(sourcePath, destinationPath, True) ' Overwrite if exists

            ' Save path and metadata to the database
            Using conn As New OleDb.OleDbConnection(connectionString)
                conn.Open()
                Dim cmd As New OleDb.OleDbCommand("INSERT INTO HouseholdDocument (HouseholdID, Title, Notes, Category, FilePath, UploadedBy, UploadDate) VALUES (?, ?, ?, ?, ?, ?, ?)", conn)
                cmd.Parameters.AddWithValue("?", 1)
                cmd.Parameters.AddWithValue("?", TextBox1.Text)
                cmd.Parameters.AddWithValue("?", TextBox2.Text)
                cmd.Parameters.AddWithValue("?", ComboBox1.Text)
                cmd.Parameters.AddWithValue("?", destinationPath) ' Save full UNC path
                cmd.Parameters.AddWithValue("?", ComboBox3.Text)
                cmd.Parameters.AddWithValue("?", DateTime.Now)
                cmd.ExecuteNonQuery()
            End Using

            MessageBox.Show("Document uploaded and saved successfully.")
            LoadDocuments()
        End If
    End Sub

    Public Sub PopulateComboboxFromDatabase(ByRef comboBox As ComboBox)
        Dim conn As New OleDbConnection(connectionString)
        Try
            Debug.WriteLine("populate combobox successful")
            'open the database connection
            conn.Open()
            'retrieve the firstname and surname columns from the Login tabel
            Dim query As String = "SELECT userName FROM Login"
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()
            'bind the retrieved data to the combobox
            ComboBox3.Items.Clear()
            While reader.Read()
                ComboBox3.Items.Add($"{reader("userName")}")
            End While
            'close the database
            reader.Close()
        Catch ex As Exception
            'handle any exeptions that may occur  
            Debug.WriteLine("failed to populate combobox")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error: {ex.StackTrace}")
        Finally
            'close the database connection
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Public Sub LoadHouseholdDocumentDatafromDatabase()

        Try

            Using conn As New OleDbConnection(connectionString)

                'conn.Open()

                'Update the table name if neccessary

                Dim tablename As String = "HouseholdDocument"

                'Create an OleDbCommand to select the data from the database

                Dim cmd As New OleDbCommand($"SELECT*FROM  {tablename}", conn)

                'create a DataAdapter and fill a DataTable

                Dim da As New OleDbDataAdapter(cmd)

                Dim dt As New DataTable()

                da.Fill(dt)

                'Bind the DataTable to the DataGridView

                DataGridView1.DataSource = dt

            End Using

        Catch ex As OleDbException

            'MessageBox.Show("$Error loading PersonalDetails data from database: {ex.message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            MessageBox.Show("$Error Loading HouseholdDocument to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch ex As Exception

            'MessageBox.Show("$unexpected Error:  {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            MessageBox.Show("$unexpected Error:" & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    'Private Sub LoadHouseholdDocument()
    '    Try
    '        Using conn As New OleDbConnection(connectionString)
    '            Dim tablename As String = "HouseholdDocument"
    '            Dim cmd As New OleDbCommand($"SELECT * FROM {tablename}", conn)
    '            Dim da As New OleDbDataAdapter(cmd)
    '            Dim dt As New DataTable()
    '            da.Fill(dt)
    '            DataGridView1.DataSource = dt

    '            ' Apply color formatting based on Category column
    '            For Each row As DataGridViewRow In DataGridView1.Rows
    '                If Not row.IsNewRow Then
    '                    Dim category As String = row.Cells("Category").Value.ToString().ToLower()
    '                    Select Case category
    '                        Case "bills"
    '                            row.DefaultCellStyle.BackColor = Color.LightBlue
    '                        Case "medical"
    '                            row.DefaultCellStyle.BackColor = Color.LightPink
    '                        Case "school"
    '                            row.DefaultCellStyle.BackColor = Color.LightGreen
    '                        Case "finance"
    '                            row.DefaultCellStyle.BackColor = Color.Khaki
    '                        Case "insurance"
    '                            row.DefaultCellStyle.BackColor = Color.Orange
    '                        Case "work"
    '                            row.DefaultCellStyle.BackColor = Color.LightGray
    '                        Case "mics", "misc", "miscellaneous"
    '                            row.DefaultCellStyle.BackColor = Color.Plum
    '                        Case Else
    '                            row.DefaultCellStyle.BackColor = Color.White
    '                    End Select
    '                End If
    '            Next

    '        End Using

    '    Catch ex As OleDbException
    '        MessageBox.Show("Error Loading HouseholdDocument to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    Catch ex As Exception
    '        MessageBox.Show("Unexpected Error: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub
    Private Sub HighlightRowsByCategory()
        For Each row As DataGridViewRow In DataGridView1.Rows
            If Not row.IsNewRow Then
                Dim category As String = row.Cells("Category").Value.ToString().ToLower()
                Select Case category
                    Case "bills"
                        row.DefaultCellStyle.BackColor = Color.LightBlue
                    Case "medical"
                        row.DefaultCellStyle.BackColor = Color.LightPink
                    Case "school"
                        row.DefaultCellStyle.BackColor = Color.LightGreen
                    Case "finance"
                        row.DefaultCellStyle.BackColor = Color.Khaki
                    Case "insurance"
                        row.DefaultCellStyle.BackColor = Color.Orange
                    Case "work"
                        row.DefaultCellStyle.BackColor = Color.LightGray
                    Case "mics", "misc", "miscellaneous"
                        row.DefaultCellStyle.BackColor = Color.Plum
                    Case Else
                        row.DefaultCellStyle.BackColor = Color.White
                End Select
            End If
        Next
    End Sub
    Private Sub LoadDocuments()
        Dim filter = ComboBox2.Text
        Dim search = TextBox4.Text
        Dim sql = "SELECT * FROM HouseholdDocument WHERE HouseholdID = ?"
        If filter <> "All" Then sql &= " AND Category = ?"
        If search <> "" Then sql &= " AND (Title LIKE ? OR Notes LIKE ?)"
        conn.Open()

        Using conn As New OleDbConnection(connectionString),
            cmd As New OleDbCommand(sql, conn)
            cmd.Parameters.AddWithValue("?", 1)
            If filter <> "All" Then cmd.Parameters.AddWithValue("?", filter)
            If search <> "" Then
                cmd.Parameters.AddWithValue("?", "%" & search & "%")
                cmd.Parameters.AddWithValue("?", "%" & search & "%")
            End If
            Dim dt As New DataTable()
            Dim adapter As New OleDbDataAdapter(cmd)
            adapter.Fill(dt)
            DataGridView1.DataSource = dt
        End Using
        conn.Close()
    End Sub

    Private Sub Household_Document_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckDatabaseConnection(StatusLabel)
        'LoadHouseholdDocument()
        'LoadFilteredDocuments()
        LoadDocuments()

        'LoadhouseholddocumentDataFromDatabase()
        LoadHouseholdDocumentDatafromDatabase()
        PopulateComboboxFromDatabase(ComboBox3)
        'ViewDocument()
        ToolTip1.SetToolTip(Button2, "Upload")
        ToolTip1.SetToolTip(Button1, "Open")
        ToolTip1.SetToolTip(Button3, "Delete")
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try
            If DataGridView1.CurrentRow IsNot Nothing Then
                Dim filePath As String = DataGridView1.CurrentRow.Cells("FilePath").Value.ToString()

                If System.IO.File.Exists(filePath) Then
                    Try
                        Process.Start(filePath)
                    Catch ex As Exception
                        MessageBox.Show("Error opening file: " & ex.Message)
                    End Try
                Else
                    MessageBox.Show("File not found: " & filePath)
                End If
            Else
                MessageBox.Show("Please select a document to open.")
            End If



        Catch ex As Exception
            'Debug.WriteLine("error selection data in the database")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("Error opening document to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try



    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        ' Check if there are any selected rows in the DataGridView for PersonalDetails  
        If DataGridView1.SelectedRows.Count > 0 Then
            ' Get the selected row  
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Assuming there is an "ID" column which is the primary key in the database  
            Dim HouseholdID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  

            ' Confirm deletion  
            Dim confirmationResult As DialogResult = MessageBox.Show("Are you sure you want to delete this document?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Hand)
            If confirmationResult = DialogResult.Yes Then
                ' Proceed with deletion  
                Try
                    Using conn As New OleDbConnection(connectionString)
                        conn.Open()

                        ' Create the delete command  
                        Dim cmd As New OleDbCommand("DELETE FROM [HouseholdDocument] WHERE [ID] = ?", conn)
                        cmd.Parameters.AddWithValue("@ID", HouseholdID) ' Primary key for matching record  

                        ' Execute the delete command  
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show("document deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ' Optionally refresh DataGridView or reload from database  

                        Else
                            MessageBox.Show("No document deleted. Please check if the ID exists.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Hand)
                        End If
                    End Using

                Catch ex As Exception
                    MessageBox.Show($"An error occurred while deleting document: {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                'LoadhouseholddocumentDataFromDatabase()
                'LoadDocuments()
            End If

        End If
    End Sub
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Try

            If DataGridView1.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim filePath As String = selectedRow.Cells("FilePath").Value.ToString() ' Adjust to your actual column name

                If File.Exists(filePath) Then
                    Dim extension As String = Path.GetExtension(filePath).ToLower()

                    If extension = ".jpg" OrElse extension = ".png" OrElse extension = ".bmp" OrElse extension = ".gif" Then
                        ' Show image in PictureBox
                        PictureBoxPreview.Image = Image.FromFile(filePath)
                        PictureBoxPreview.SizeMode = PictureBoxSizeMode.Zoom

                    ElseIf extension = ".pdf" Then
                        ' For PDF, you may need to use a PDF viewer control or load the PDF in a WebBrowser
                        WebBrowser1.Navigate(filePath) ' Add a WebBrowser control named WebBrowser1
                        PictureBoxPreview.Image = Nothing
                    Else
                        MessageBox.Show("Unsupported file type.")
                        PictureBoxPreview.Image = Nothing
                    End If
                Else
                    MessageBox.Show("File not found.")
                End If
            End If

            Debug.WriteLine("selecting data in the datagridview")
            If DataGridView1.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

                ' Load the data from the selected row into UI controls  
                ComboBox1.Text = selectedRow.Cells("Category").Value.ToString()
                TextBox1.Text = selectedRow.Cells("Title").Value.ToString()
                TextBox2.Text = selectedRow.Cells("Notes").Value.ToString()
                TextBox3.Text = selectedRow.Cells("FilePath").Value.ToString()
                DateTimePicker1.Text = selectedRow.Cells("UploadDate").Value.ToString()
                ComboBox3.Text = selectedRow.Cells("UploadedBy").Value.ToString()

            End If


        Catch ex As Exception
            Debug.WriteLine("error selection data in the database")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            'MessageBox.Show("Error saving document to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        PrintDialog1.Document = PrintDocument1

        If PrintDialog1.ShowDialog() = DialogResult.OK Then

            LoadFilteredDocuments() ' Load filtered data based on selected Category

            If DataGridView1.Rows.Count > 0 Then

                PrintDocument1.Print()

            Else

                MessageBox.Show("No document found for the selected period.", "Print Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            End If

        End If

    End Sub


    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        Dim dv As DataView = CType(DataGridView1.DataSource, DataTable).DefaultView
        dv.RowFilter = $"Title LIKE '%{TextBox4.Text}%' OR Notes LIKE '%{TextBox4.Text}%'"

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim dv As DataView = CType(DataGridView1.DataSource, DataTable).DefaultView
        dv.RowFilter = $"Category = '{ComboBox2.Text}'"
        ''ApplySearchAndFilter()
        ''LoadFilteredDocuments()

    End Sub


    Private DataGridView As DataTable

    Private currentRowIndex As Integer = 0

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage


        Dim font As New Font("Arial", 12)

        Dim brush As New SolidBrush(Color.Black)

        Dim yPos As Integer = 100

        Dim leftMargin As Integer = e.MarginBounds.Left

        ' Print Title

        e.Graphics.DrawString("document Report", New Font("Arial", 16, FontStyle.Bold), brush, leftMargin, yPos)

        yPos += 40

        ' Check if there is data to print

        If DataGridView1.Rows.Count = 0 Then

            e.Graphics.DrawString("No data available.", font, brush, leftMargin, yPos)

            Exit Sub

        End If

        ' Print filtered meal plan data

        For Each row As DataRow In DataGridView1.Rows

            e.Graphics.DrawString("Title: " & row("Title").ToString(), font, brush, leftMargin, yPos)

            yPos += 30

            e.Graphics.DrawString("Notes: " & row("Notes").ToShortDateString(), font, brush, leftMargin, yPos)

            yPos += 30
            e.Graphics.DrawString("Category: " & row("Category").ToString(), font, brush, leftMargin, yPos)

            yPos += 30
            e.Graphics.DrawString("FilePath: " & row("FilePath").ToString(), font, brush, leftMargin, yPos)

            yPos += 30

            e.Graphics.DrawString("UploadedBy: " & row("UploadedBy").ToString(), font, brush, leftMargin, yPos)

            yPos += 40

            e.Graphics.DrawString("UploadDate: " & Convert.ToDateTime(row("UploadDate")).ToShortDateString(), font, brush, leftMargin, yPos)
            yPos += 40

        Next

    End Sub


    ' Load filtered document data based on category
    Private Sub LoadFilteredDocuments()
        Using dbConnection As New OleDbConnection(connectionString)
            Dim selectedCategory As String = ComboBox3.SelectedItem?.ToString()
            Dim query As String = "SELECT * FROM HouseholdDocument WHERE 1=1"

            ' Add category filter if selected
            If Not String.IsNullOrEmpty(selectedCategory) Then
                query &= " AND Category = ?"
            End If

            Dim cmd As New OleDbCommand(query, dbConnection)

            ' Bind parameters in the same order as in the query
            If Not String.IsNullOrEmpty(selectedCategory) Then
                cmd.Parameters.AddWithValue("?", selectedCategory)
            End If

            Dim adapter As New OleDbDataAdapter(cmd)
            DataGridView = New DataTable()
            adapter.Fill(DataGridView)

            DataGridView1.DataSource = DataGridView ' Display filtered data in DataGridView
        End Using
    End Sub

End Class