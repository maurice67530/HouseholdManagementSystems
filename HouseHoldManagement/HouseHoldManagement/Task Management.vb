Imports System.Data.OleDb

Public Class Task_Management
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\khodani\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"
    Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
    Private Sub Task_Management_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim tooltip As New ToolTip
        tooltip.SetToolTip(Button1, "Submit")
        tooltip.SetToolTip(Button4, "Refresh")
        tooltip.SetToolTip(Button3, "Delete")
        tooltip.SetToolTip(Button2, "Edit")
        tooltip.SetToolTip(Button8, "Dashboard")
        tooltip.SetToolTip(Button5, "Filter")
        tooltip.SetToolTip(Button6, "Sort")

        ComboBox1.Items.AddRange(New String() {"Low", "Medium", "High"})
        ComboBox2.Items.AddRange(New String() {"Not started", "In progress", "Completed"})
    End Sub
    Private Sub ComboBox3_click(sender As Object, e As EventArgs) Handles ComboBox3.Click
        PopulateComboboxFromDatabase(ComboBox3)
    End Sub
    Public Sub PopulateComboboxFromDatabase(ByRef comboBox As ComboBox)
        Dim conn As New OleDbConnection(HouseHoldManagment_Module.connectionString)
        Try
            Debug.WriteLine("populating combobox from database successfully!")
            ' 1. Open the database connection  
            conn.Open()

            ' 2. Retrieve the FirstName and LastName columns from the PersonalDetails table  
            Dim query As String = "SELECT FirstName, LastName FROM Person"
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader()

            ' 3. Bind the retrieved data to the combobox  
            comboBox.Items.Clear()
            While reader.Read()
                comboBox.Items.Add($"{reader("FirstName")} {reader("LastName")}")
            End While

            ' 4. Close the database connection  
            reader.Close()

        Catch ex As Exception
            Debug.WriteLine($"form loaded unsuccessful")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
        Finally
            ' Close the database connection  
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Debug.WriteLine("entering btnsave")

            If TextBox2.Text = " " Then
                MsgBox("please enter a value")
                TextBox2.Focus()

            End If
            If ComboBox1.Text = "" Then
                MsgBox("please enter a value")
                ComboBox1.Focus()

            End If
            If DateTimePicker1.Text = "" Then
                MsgBox("please enter a value")
                DateTimePicker1.Focus()

            End If
            Using conn As New OleDbConnection(connectionString)
                conn.Open()

                Dim tableName As String = "Tasks"
                Dim cmd As New OleDbCommand("INSERT INTO Tasks ([Title],[Description],[DueDate],[Priority],[Status],[Assignedto]) VALUES (@Title, @Description, @DueDate, @Priority, @Status, @AssignedTo)", conn)

                Dim Task As New DailyTask() With {
             .Title = TextBox1.Text,
              .Description = TextBox2.Text,
            .DueDate = DateTimePicker1.Text,
            .Priority = ComboBox1.Text,
            .Status = ComboBox2.Text,
            .AssignedTo = ComboBox3.Text}


                'parameters

                cmd.Parameters.AddWithValue("@Title", Task.Title)
                cmd.Parameters.AddWithValue("@Description", Task.Description)
                cmd.Parameters.AddWithValue("@DueDate", Task.DueDate)
                cmd.Parameters.AddWithValue("@Priority", Task.Priority)
                cmd.Parameters.AddWithValue("@Status", Task.Status)
                cmd.Parameters.AddWithValue("@AssignedTo", Task.AssignedTo)

                cmd.ExecuteNonQuery()

                MsgBox("tasks Added!" & vbCrLf &
                   "Title: " & Task.Title.ToString() & vbCrLf &
                   "Description: " & Task.Description & vbCrLf &
                   "DueDate: " & Task.DueDate & vbCrLf &
                   "Priority: " & Task.Priority & vbCrLf &
                   "Status: " & Task.Status & vbCrLf &
                   "Assignedto: " & Task.AssignedTo, vbInformation, "Tasks Confirmation")


            End Using
        Catch ex As OleDbException
            Debug.WriteLine($"database error: {ex.Message}")
            MessageBox.Show("error saving test to database. please check the connection.", "Database error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch ex As Exception
            Debug.WriteLine($"general error in btnsave:{ex.Message}")
            MessageBox.Show("An unexpected error occurred.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
End Class