Imports System.Data.OleDb
Imports System.IO

Public Class Household_Document
    Public Property conn As New OleDbConnection(connectionString)

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim ofd As New OpenFileDialog()
        ofd.Filter = "Documents|*.pdf;*.docx;*.xlsx;*.jpg;*.png|All files|*.*"
        If ofd.ShowDialog() = DialogResult.OK Then
            Dim sourcePath = ofd.FileName
            Dim fileName = IO.Path.GetFileName(sourcePath)
            Dim categoryFolder = Path.Combine("C:\DocumentLibrary", ComboBox1.Text)
            Directory.CreateDirectory(categoryFolder)
            Dim destPath = Path.Combine(categoryFolder, fileName)
            File.Copy(sourcePath, destPath, True)

            Using conn As New OleDbConnection(connectionString)
                Dim cmd As New OleDbCommand("INSERT INTO HouseholdDocument (Title, Notes, Category, FilePath, UploadedBy, UploadDate)
            VALUES (@Title, @Notes, @Category, @FilePath, @UploadedBy, @UploadDate)", conn)
                'cmd.Parameters.AddWithValue("@HouseholdID", ID)
                cmd.Parameters.AddWithValue("@Title", TextBox1.Text)
                cmd.Parameters.AddWithValue("@Notes", TextBox2.Text)
                cmd.Parameters.AddWithValue("@Category", ComboBox1.Text)
                cmd.Parameters.AddWithValue("@FilePath", destPath)
                cmd.Parameters.AddWithValue("@UploadedBy", Environment.UserName)
                cmd.Parameters.AddWithValue("@UploadDate", DateTime.Now)
                conn.Open()
                cmd.ExecuteNonQuery()
            End Using

            MessageBox.Show("Document uploaded.")
            LoadDocuments()
        End If

    End Sub

    Public Sub LoadhouseholddocumentDataFromDatabase()
        Try
            Debug.WriteLine("DataGridview populated successfully ChoresForm_Load")
            Using conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
                conn.Open()

                ' Update the table name if necessary  
                Dim tableName As String = "HouseholdDocument"


                ' Create an OleDbCommand to select the data from the database  
                Dim cmd As New OleDbCommand($"SELECT * FROM {tableName}", conn)

                ' Create a DataAdapter and fill a DataTable  
                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable()
                da.Fill(dt)

                ' Bind the DataTable to the DataGridView  
                DataGridView1.DataSource = dt
            End Using

        Catch ex As Exception
            Debug.WriteLine($"DataGridView population failed")
            Debug.WriteLine($"Unexpected error in DataGridView: {ex.Message}")
            Debug.WriteLine($"Error in PopulateDataGridView: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            'MessageBox.Show("An error occurred while loading data into the grid.", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub LoadDocuments()
        If ComboBox1.SelectedIndex = -1 Then Exit Sub

        Dim conn As New OleDbConnection("HouseHoldManagment_Module.connectionstring")
        Dim cmd As New OleDbCommand("SELECT SelectHouseHold, FileName, FilePath, UploadedDate FROM HouseholdDocuments WHERE ID= @ID", conn)
        cmd.Parameters.AddWithValue("@SelectHouseHold", ComboBox1.SelectedValue)

        Dim adapter As New OleDbDataAdapter(cmd)
        Dim dt As New DataTable()
        adapter.Fill(dt)

        DataGridView1.DataSource = dt
        DataGridView1.Columns("FilePath").Visible = False ' Optional
    End Sub
    Private Sub Household_Document_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'LoadFilteredDocuments()


        LoadhouseholddocumentDataFromDatabase()
        LoadDocuments()
        'ViewDocument()
        ToolTip1.SetToolTip(Button2, "Upload")
        ToolTip1.SetToolTip(Button1, "Open")
        ToolTip1.SetToolTip(Button3, "Delete")
    End Sub

    'Private Sub LoadFilteredDocuments()
    '    ListBox1.Items.Clear()

    '    Dim conn As New OleDbConnection(connectionString)
    '    Dim query As String = "SELECT Title FROM HouseholdDocument WHERE HouseholdID = @hid"

    '    If ComboBox2.Text <> "All" Then
    '        query &= " AND Category = @cat"
    '    End If

    '    Dim cmd As New OleDbCommand(query, conn)
    '    'cmd.Parameters.AddWithValue("?", HouseholdID)
    '    If ComboBox2.Text <> "All" Then
    '        cmd.Parameters.AddWithValue("@cat", ComboBox2.Text)
    '    End If

    '    conn.Open()
    '    Dim reader As OleDbDataReader = cmd.ExecuteReader()
    '    While reader.Read()
    '        ListBox1.Items.Add(reader("Title").ToString())
    '    End While
    '    conn.Close()
    'End Sub
    Private Sub ViewDocument(SelectHouseHold As String)
        Dim connStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\khodani\Documents\yourdb.accdb"
        Dim filePath As String = ""

        Using conn As New OleDbConnection(connStr)
            conn.Open()
            Dim cmd As New OleDbCommand("SELECT FilePath FROM HouseholdDocument WHERE ID = @ID", conn)
            cmd.Parameters.AddWithValue("@ID", SelectHouseHold)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()
            If reader.Read() Then
                filePath = reader("FilePath").ToString()
            End If
        End Using

        If System.IO.File.Exists(filePath) Then
            Process.Start(filePath)
        Else
            MessageBox.Show("File not found: " & filePath)
        End If
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
                LoadhouseholddocumentDataFromDatabase()
            End If

        End If
    End Sub
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Try

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
            MessageBox.Show("Error saving document to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim pd As New PrintDialog()
        ' Logic to generate and send document list to printer
        ' Or export to PDF using iTextSharp or similar

    End Sub


    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        Dim dv As DataView = CType(DataGridView1.DataSource, DataTable).DefaultView
        dv.RowFilter = $"Title LIKE '%{TextBox4.Text}%' OR Notes LIKE '%{TextBox4.Text}%'"

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim dv As DataView = CType(DataGridView1.DataSource, DataTable).DefaultView
        dv.RowFilter = $"Category = '{ComboBox2.Text}'"
        'ApplySearchAndFilter()
        'LoadFilteredDocuments()
    End Sub


End Class