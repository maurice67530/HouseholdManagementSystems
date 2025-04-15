Imports System.IO
Imports System.Data.OleDb
Public Class GroceryItemss
    Dim conn As New OleDbConnection(connectionString)
    Public Const connectionString As String = " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Rotondwa\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb;Persist Security Info=False;"
    Private Sub GroceryItemss_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Start()
        LoadGrocerydataFromDatabase()

    End Sub
    Public Sub LoadGrocerydataFromDatabase()
        Try
            Debug.WriteLine("Populate Datagridview: Datagridview populated successfully.")
            Using conn As New OleDbConnection(connectionString)
                conn.Open()

                Dim tableName As String = "GroceryItem"

                Dim cmd As New OleDbCommand($"SELECT*FROM {tableName}", conn)

                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)
                DataGridView1.DataSource = dt
            End Using
        Catch ex As Exception
            Debug.WriteLine("Failed to populate Gatagridview")
            Debug.Write($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error Loading grocery data from database: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ToolTip1.SetToolTip(Button5, "HighLight")
        ToolTip1.SetToolTip(Button4, "DashBoard")
        ToolTip1.SetToolTip(Button3, "Clear")
        ToolTip1.SetToolTip(Button1, "Submit")
        ToolTip1.SetToolTip(Button2, "Update")


    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.Cells("ExpiryDate").Value IsNot Nothing Then
                    Dim ExpiryDate As DateTime = Convert.ToDateTime(row.Cells("ExpiryDate").Value)
                    Dim ItemName As String = row.Cells("ItemName").Value.ToString

                    If ExpiryDate < DateTime.Now AndAlso ItemName <> "" Then
                        row.DefaultCellStyle.BackColor = Color.HotPink

                    Else
                        row.DefaultCellStyle.BackColor = Color.White
                    End If
                End If
            Next
            Dim incompleteCount As Integer = 0
            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.Cells("ItemName").Value IsNot Nothing AndAlso row.Cells("ItemName").Value.ToString() <> "" Then
                    incompleteCount += 1

                End If
            Next
            Label5.Text = "Expired Grocery:" & incompleteCount.ToString
        Catch ex As Exception
            MessageBox.Show("error highlighting expired grocery")
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Debug.WriteLine($"Entering btnDelete_Click")


        If DataGridView1.SelectedRows.Count > 0 Then
            ' Get the selected row  
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Assuming there is an "ID" column which is the primary key in the database  
            Dim expenseId As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Replace "ID" with your actual column name  

            ' Confirm deletion  
            Dim confirmationResult As DialogResult = MessageBox.Show("Are you sure you want to delete this Grocery?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If confirmationResult = DialogResult.Yes Then
                ' Proceed with deletion  
                Try
                    Using conn As New OleDbConnection(connectionString)
                        conn.Open()

                        ' Create the delete command  
                        Dim cmd As New OleDbCommand("DELETE FROM [GroceryItem] WHERE [ID] = ?", conn)
                        cmd.Parameters.AddWithValue("@ID", expenseId) ' Primary key for matching record  

                        ' Execute the delete command  
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                        If rowsAffected > 0 Then
                            MessageBox.Show($"Grocery deleted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ' Optionally refresh DataGridView or reload from database  
                            'PopulateComboBox()
                        Else
                            MessageBox.Show($"No Grocery was deleted. Please check if the ID exists.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End Using

                Catch ex As Exception
                    MessageBox.Show($"An error occurred while deleting the expense: {ex.Message}", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            MessageBox.Show($"Please select an Inventoty to delete.", "Deletion Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

        Debug.WriteLine($"Entering btnDelete_Click")

        Debug.WriteLine($"Exiting btnDelete_Click")
        LoadGrocerydataFromDatabase()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
        Dashboard.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Using conn As New OleDbConnection(connectionString)

                conn.Open()
                Dim cmd As New OleDbCommand($"INSERT INTO GroceryItem ([ItemName], [Quantity], [Unit], [ExpiryDate], [Category], [PricePerUnit], [IsPurchased] ) VALUES (?, ?, ?, ?, ?, ?, ?)", conn)

                cmd.Parameters.Clear()
                cmd.Parameters.AddWithValue("@ItemName", TextBox1.Text)
                cmd.Parameters.AddWithValue("@Quantity", TextBox2.Text)
                cmd.Parameters.AddWithValue("@Unit", TextBox3.Text)
                cmd.Parameters.AddWithValue("@ExpiryDate", DateTimePicker1.Value)
                cmd.Parameters.AddWithValue("@Category", TextBox4.Text)
                cmd.Parameters.AddWithValue("@PricePerUnit", TextBox5.Text)
                cmd.Parameters.AddWithValue("@IsPurchased", TextBox6.Text)


                cmd.ExecuteNonQuery()
                conn.Close()
                LoadGrocerydataFromDatabase()
                MsgBox("Grocery Added!" & vbCrLf &
                       "ItemName: " & TextBox1.Text & vbCrLf &
                        "Quantity: " & TextBox2.Text & vbCrLf &
                        "Unit: " & TextBox3.Text & vbCrLf &
                        "ExpiryDate: " & DateTimePicker1.Text & vbCrLf &
                        "Category: " & TextBox4.Text & vbCrLf &
                        "PricePerUnit: " & TextBox5.Text & vbCrLf &
                        "IsPurchased: " & TextBox6.Text)


            End Using
        Catch ex As OleDbException
            Debug.WriteLine($"Database error: {ex.Message}")
            Debug.Write($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show("Error saving grocery to database: Please check the connectivity." & ex.Message & vbNewLine & ex.StackTrace, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Debug.WriteLine($"General error: {ex.Message}")

            MessageBox.Show("Unexpected Error: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Debug.WriteLine("Entering button update click")
        If DataGridView1.SelectedRows.Count = 0 Then
            Debug.WriteLine("User confirmed update")
            MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Debug.WriteLine($"Format error in button update:")
            Debug.WriteLine("Updating data: data updated")
            Dim ItemName As String = TextBox1.Text
            Dim Quantity As String = TextBox2.Text
            Dim Unit As String = TextBox3.Text
            Dim ExpiryDate As String = DateTimePicker1.Value
            Dim Category As String = TextBox4.Text
            Dim PricePerUnit As String = TextBox5.Text
            Dim IsPurchased As String = TextBox6.Text

            Using conn As New OleDbConnection(connectionString)
                conn.Open()

                ' Get the ID of the selected row (assuming your table has a primary key named "ID")  
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim ID As Integer = Convert.ToInt32(selectedRow.Cells("ID").Value) ' Change "ID" to your primary key column name  

                ' Create an OleDbCommand to update the personnel data in the database  
                Dim cmd As New OleDbCommand("UPDATE [GroceryItem] SET [ItemName] = ?, [Quantity] = ?, [Unit] = ?, [ExpiryDate] = ?, [Category] = ?, [PricePerUnit] = ?, [IsPurchased] = ? WHERE [ID] = ?", conn)

                ' Set the parameter values from the UI controls  

                'cmd.Parameters.AddWithValue("@PhotoID", PhotoID.textbox1.Text)
                cmd.Parameters.AddWithValue("@ItemName", ItemName)
                cmd.Parameters.AddWithValue("@Quantity", Quantity)
                cmd.Parameters.AddWithValue("@Unit", Unit)
                cmd.Parameters.AddWithValue("@ExpiryDate", ExpiryDate)
                cmd.Parameters.AddWithValue("@Category", Category)
                cmd.Parameters.AddWithValue("@PricePerUnit", PricePerUnit)
                cmd.Parameters.AddWithValue("@IsPurchased", IsPurchased)
                cmd.Parameters.AddWithValue("@ID", ID)

                cmd.ExecuteNonQuery()

                MsgBox("Grocery Items Updated Successfuly!", vbInformation, "Update Confirmation")
                LoadGrocerydataFromDatabase()
                HouseHoldManagement.ClearControls(Me)
            End Using
        Catch ex As OleDbException
            Debug.WriteLine("User cancelled update")
            Debug.WriteLine("Unexpected error in button update")
            Debug.Write($"Stack Trace: {ex.StackTrace}")
            MessageBox.Show($"Error updating grocery in database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show($"Unexpected error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Debug.WriteLine("Exiting button update")
    End Sub
    Public Sub notify()
        Try
            Using conn As New OleDbConnection(connectionString)
                ' Open connection to the database
                conn.Open()

                ' SQL Query to get tasks that are not started
                Dim query As String = "SELECT ItemName, ExpiryDate FROM GroceryItem "
                Dim cmd As New OleDbCommand(query, conn)

                ' Execute the query and get the results
                Dim reader As OleDbDataReader = cmd.ExecuteReader()

                ' If there are any tasks not started
                If reader.HasRows Then
                    ' Prepare the email body
                    Dim Notification As String = "The following Grocery has Expired:" & Environment.NewLine

                    While reader.Read()
                        Notification &= "GroceryItem: " & reader("ItemName").ToString() & ", By: " & reader("ExpiryDate").ToString() & Environment.NewLine
                    End While
                    'Timer1.Stop() : MessageBox.Show("The following Expense has ecxeeded 80%.", "Expense Alert", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Timer1.Stop() : MessageBox.Show("Item Expired: " & vbCrLf & Notification, "Tasks Alert", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                End If
            End Using
        Catch ex As Exception
            ' Handle any errors (e.g., connection errors, query errors)
            Console.WriteLine("Error: " & ex.Message)
        Finally
            ' Close the connection
            'conn.Close()
        End Try


    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1 = New Timer
        Timer1.Interval = 20
        Timer1.Enabled = True
        notify() ' Check every 1 minute (60000 milliseconds)
    End Sub
End Class