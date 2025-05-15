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
                Dim cmd As New OleDbCommand("INSERT INTO HouseholdDocument (HouseholdID, Title, Notes, Category, FilePath, UploadedBy, UploadDate)
            VALUES (@HouseholdID, @Title, @Notes, @Category, @FilePath, @UploadedBy, @UploadDate)", conn)
                cmd.Parameters.AddWithValue("@HouseholdID", 1)
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
        Dim filter = ComboBox2.Text
        Dim search = TextBox4.Text
        Dim sql = "SELECT * FROM HouseholdDocument WHERE HouseholdID = ?"
        If filter <> "All" Then sql &= " AND Category = ?"
        If search <> "" Then sql &= " AND (Title LIKE ? OR Notes LIKE ?)"
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
    End Sub
    'Private Sub LoadDocuments()
    '    If ComboBox1.SelectedIndex = -1 Then Exit Sub

    '    Dim conn As New OleDbConnection("HouseHoldManagment_Module.connectionstring")
    '    Dim cmd As New OleDbCommand("SELECT SelectHouseHold, FileName, FilePath, UploadedDate FROM HouseholdDocuments WHERE ID= @ID", conn)
    '    cmd.Parameters.AddWithValue("@SelectHouseHold", ComboBox1.SelectedValue)

    '    Dim adapter As New OleDbDataAdapter(cmd)
    '    Dim dt As New DataTable()
    '    adapter.Fill(dt)

    '    DataGridView1.DataSource = dt
    '    DataGridView1.Columns("FilePath").Visible = False ' Optional
    'End Sub
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
    'Private Sub ViewDocument(SelectHouseHold As String)
    '    Dim connStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\khodani\Documents\yourdb.accdb"
    '    Dim filePath As String = ""

    '    Using conn As New OleDbConnection(connStr)
    '        conn.Open()
    '        Dim cmd As New OleDbCommand("SELECT FilePath FROM HouseholdDocument WHERE ID = @ID", conn)
    '        cmd.Parameters.AddWithValue("@ID", SelectHouseHold)
    '        Dim reader As OleDbDataReader = cmd.ExecuteReader()
    '        If reader.Read() Then
    '            filePath = reader("FilePath").ToString()
    '        End If
    '    End Using

    '    If System.IO.File.Exists(filePath) Then
    '        Process.Start(filePath)
    '    Else
    '        MessageBox.Show("File not found: " & filePath)
    '    End If
    'End Sub
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
        'Dim dv As DataView = CType(DataGridView1.DataSource, DataTable).DefaultView
        'dv.RowFilter = $"Category = '{ComboBox2.Text}'"
        ''ApplySearchAndFilter()
        ''LoadFilteredDocuments()

    End Sub

    'Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
    '    ' Get selected values from ComboBoxes
    '    Dim selectedCategory As String = If(ComboBox1.SelectedItem IsNot Nothing, ComboBox2.SelectedItem.ToString(), "")
    '    Dim selectedUploadedBy As String = If(ComboBox3.SelectedItem IsNot Nothing, ComboBox1.SelectedItem.ToString(), "")

    '    ' Call the filtering method with selected values
    '    Filterdocument(selectedCategory, selectedUploadedBy)
    'End Sub
    'Public Sub Filterdocument(Category As String, UploadedBy As String)
    '    Dim taskTable As New DataTable()
    '    Try
    '        conn.Open()
    '        Dim query As String = "SELECT * FROM HouseholdDocument WHERE 1=1"

    '        ' Only add conditions if filters are selected  
    '        If Not String.IsNullOrEmpty(Category) Then
    '            query &= " AND Category = @Category"
    '        End If

    '        If Not String.IsNullOrEmpty(UploadedBy) Then
    '            query &= " AND UploadedBy = @UploadedBy"
    '        End If

    '        Dim command As New OleDb.OleDbCommand(query, conn)

    '        ' Only add parameters if filters are selected  
    '        If Not String.IsNullOrEmpty(Category) Then
    '            command.Parameters.AddWithValue("@UploadedBy", UploadedBy)
    '        End If

    '        If Not String.IsNullOrEmpty(UploadedBy) Then
    '            command.Parameters.AddWithValue("@Category", Category)
    '        End If

    '        Dim adapter As New OleDb.OleDbDataAdapter(command)
    '        adapter.Fill(taskTable)
    '        DataGridView1.DataSource = taskTable
    '    Catch ex As Exception
    '        MsgBox("Error filtering document: " & ex.Message, MsgBoxStyle.Critical, "Database Error")
    '    Finally
    '        conn.Close()
    '    End Try
    '    'Daily_task.loadiTaskmanagementfromdatabase()
    'End Sub
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

            e.Graphics.DrawString("Notes: " & Convert.ToDateTime(row("Notes")).ToShortDateString(), font, brush, leftMargin, yPos)

            yPos += 30
            e.Graphics.DrawString("Category: " & row("Category").ToString(), font, brush, leftMargin, yPos)

            yPos += 30
            e.Graphics.DrawString("FilePath: " & Convert.ToDateTime(row("FilePath")).ToShortDateString(), font, brush, leftMargin, yPos)

            yPos += 30

            e.Graphics.DrawString("UploadedBy: " & row("UploadedBy").ToString(), font, brush, leftMargin, yPos)

            yPos += 40

            e.Graphics.DrawString("UploadDate: " & row("UploadDate").ToString(), font, brush, leftMargin, yPos)

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