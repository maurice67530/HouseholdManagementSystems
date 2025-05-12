Imports System.Data.OleDb
Imports System.IO

Public Class Household_Document
    Public Property conn As New OleDbConnection(khodani.connectionString)
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        ' Ensure a household is selected
        If ComboBox1.SelectedIndex = -1 Then
            MessageBox.Show("Please select a household.")
            Return
        End If

        ' Open file dialog to select document
        Dim openFileDialog1 As New OpenFileDialog()
        If openFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim sourcePath As String = openFileDialog1.FileName
            Dim fileName As String = Path.GetFileName(sourcePath)
            Dim targetDirectory As String = "C:\Documents\UploadedFiles\"
            Dim targetPath As String = Path.Combine(targetDirectory, fileName)


            ' Ensure target directory exists
            If Not Directory.Exists(targetDirectory) Then
                Directory.CreateDirectory(targetDirectory)
            End If

            ' Copy the file to the target directory
            File.Copy(sourcePath, targetPath, True)

            ' Connection string to Access DB
            Dim connStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\khodani\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb;"
            Using conn As New OleDbConnection(connStr)
                conn.Open()
                Dim cmd As New OleDbCommand("INSERT INTO HouseholdDocument (SelectHouseHold, FileName, FilePath, UploadedDate, UploadedBy) " &
                                            "VALUES (@SelectHouseHold, @FileName, @FilePath, @UploadedDate, @UploadedBy)", conn)
                cmd.Parameters.AddWithValue("@SelectHouseHold", ComboBox1.SelectedItem.ToString())
                cmd.Parameters.AddWithValue("@FileName", fileName)
                cmd.Parameters.AddWithValue("@FilePath", targetPath)
                cmd.Parameters.AddWithValue("@UploadedDate", DateTime.Now)
                cmd.Parameters.AddWithValue("@UploadedBy", Environment.UserName)

                cmd.ExecuteNonQuery()
                MessageBox.Show("File uploaded and data saved successfully.")
            End Using
            LoadhouseholddocumentDataFromDatabase()
        End If
    End Sub

    Public Sub LoadhouseholddocumentDataFromDatabase()
        Try
            Debug.WriteLine("DataGridview populated successfully ChoresForm_Load")
            Using conn As New OleDbConnection(khodani.connectionString)
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

        Dim conn As New OleDbConnection("khodani.connectionstring")
        Dim cmd As New OleDbCommand("SELECT SelectHouseHold, FileName, FilePath, UploadedDate FROM HouseholdDocuments WHERE ID= @ID", conn)
        cmd.Parameters.AddWithValue("@SelectHouseHold", ComboBox1.SelectedValue)

        Dim adapter As New OleDbDataAdapter(cmd)
        Dim dt As New DataTable()
        adapter.Fill(dt)

        DataGridView1.DataSource = dt
        DataGridView1.Columns("FilePath").Visible = False ' Optional
    End Sub
    Private Sub Household_Document_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadhouseholddocumentDataFromDatabase()
        'ViewDocument()
        toolTip1.SetToolTip(Button2, "Upload")
        toolTip1.SetToolTip(Button1, "Open")
        toolTip1.SetToolTip(Button3, "Delete")
    End Sub

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



        'Dim SelectHousehold As String = ComboBox1.Text ' or get from selected row
        'ViewDocument(SelectHousehold)




        Try

            '    ' Example: retrieve file path from a selected item in a ListBox or database
            '    Dim filePath As String = "C:\Users\khodani\Source\Repos\maurice67530\HouseholdManagementSystems\HouseHoldManagement Kundani\HouseHoldManagement Kundani" ' replace this with dynamic value

            '    If File.Exists(filePath) Then
            '        Process.Start(filePath)
            '    Else
            '        MessageBox.Show("File not found: " & filePath)
            '    End If

            '    ' Get dynamic values from UI controls or environment
            '    Dim SelectHouseHold As String = ComboBox1.SelectedItem.ToString()
            '    Dim FileName As String = Path.GetFileName(OpenFileDialog1.FileName) ' assuming a file was selected
            '    'Dim FilePath As String = "C:\Uploads\" & FileName ' or your actual path logic
            '    Dim UploadedDate As DateTime = DateTime.Now
            '    Dim UploadedBy As String = Environment.UserName
            '    'Dim DocumentType As String = TextBoxDocumentType.Text ' assuming you have this TextBox on form

            '    Dim cmd As New OleDbCommand("INSERT INTO HouseholdDocument (SelectHouseHold, FileName, FilePath, UploadedDate, UploadedBy) " &
            '                                "VALUES (@SelectHouseHold, @FileName, @FilePath, @UploadedDate, @UploadedBy)", conn)

            '    cmd.Parameters.AddWithValue("@SelectHouseHold", SelectHouseHold)
            '    cmd.Parameters.AddWithValue("@FileName", FileName)
            '    cmd.Parameters.AddWithValue("@FilePath", filePath)
            '    cmd.Parameters.AddWithValue("@UploadedDate", UploadedDate)
            '    cmd.Parameters.AddWithValue("@UploadedBy", UploadedBy)

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
            Dim MealPlansId As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  

            ' Confirm deletion  
            Dim confirmationResult As DialogResult = MessageBox.Show("Are you sure you want to delete this document?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Hand)
            If confirmationResult = DialogResult.Yes Then
                ' Proceed with deletion  
                Try
                    Using conn As New OleDbConnection(khodani.connectionString)
                        conn.Open()

                        ' Create the delete command  
                        Dim cmd As New OleDbCommand("DELETE FROM [HouseholdDocument] WHERE [ID] = ?", conn)
                        cmd.Parameters.AddWithValue("@ID", MealPlansId) ' Primary key for matching record  

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
                    MessageBox.Show($"An error occurred while deleting the mealplan: {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                LoadhouseholddocumentDataFromDatabase()
            End If

        End If
    End Sub



    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Try

            Debug.WriteLine("selecting data in the datagridview")
            If DataGridView1.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

                ' Load the data from the selected row into UI controls  
                ComboBox1.Text = selectedRow.Cells("SelectHouseHold").Value.ToString()
                TextBox1.Text = selectedRow.Cells("FileName").Value.ToString()
                TextBox2.Text = selectedRow.Cells("FilePath").Value.ToString()
                DateTimePicker1.Text = selectedRow.Cells("UploadedDate").Value.ToString()
                ComboBox3.Text = selectedRow.Cells("UploadedBy").Value.ToString()

            End If
        Catch ex As Exception
            Debug.WriteLine("error selection data in the database")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("Error saving inventory to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class