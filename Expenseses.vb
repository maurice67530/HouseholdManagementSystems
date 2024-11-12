
Imports System.IO
Imports System.Windows.Forms
Imports System.Data.OleDb
Public Class Expenseses
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Using connect As New OleDbConnection(Module1.connectionString)
                connect.Open()

                ' Update the table name if necessary  
                Dim tableName As String = "Expenses"

                ' Create an OleDbCommand to insert the person data into the database  
                Dim cmd As New OleDbCommand("INSERT INTO Expenses ([Amount], [Currency], [payment], [Tags], [frequency], [approvalStatus], [Dates]) VALUES ( ?, ?, ?, ?, ?, ?, ?)", connect)

                ' Set the parameter values from the UI controls 
                'Class declaretions
                Dim Exp As New Expenses

                'Assign Values 
                Exp.Amount = TextBox1.Text
                Exp.Currency = ComboBox1.SelectedItem.ToString
                Exp.payment = ComboBox2.SelectedItem.ToString
                Exp.Tags = TextBox2.Text
                Exp.frequency = ComboBox3.SelectedItem.ToString
                Exp.approvalStatus = ComboBox4.SelectedItem.ToString
                Exp.Dates = DateTimePicker1.Value.ToString()




                'For Each person As person In Personal
                cmd.Parameters.Clear()

                'cmd.Parameters.AddWithValue("@ID",TextBox8)
                cmd.Parameters.AddWithValue("@Amount", Exp.Amount)
                cmd.Parameters.AddWithValue("@Currency", Exp.Currency)
                cmd.Parameters.AddWithValue("@payment", Exp.payment)
                cmd.Parameters.AddWithValue("@Tags", Exp.Tags)
                cmd.Parameters.AddWithValue("@frequency", Exp.frequency)
                cmd.Parameters.AddWithValue("@approvalStatus", Exp.approvalStatus)
                cmd.Parameters.AddWithValue("@Dates", Exp.Dates)


                MsgBox(" You are now added  !" & vbCrLf &
                      "Amount: " & Exp.Amount & vbCrLf &
                      "Currency: " & Exp.Currency & vbCrLf &
                      "payment:" & Exp.payment & vbCrLf &
                      "Tags: " & Exp.Tags & vbCrLf &
                      "frequency: " & Exp.frequency & vbCrLf &
                      "approvalStatus: " & Exp.approvalStatus & vbCrLf &
                      "Dates: " & Exp.Dates & vbCrLf, vbInformation, "Credentials  confirmation")

                MessageBox.Show("Testing information saved to Database successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' Execute the SQL command to insert the data  
                cmd.ExecuteNonQuery()
                'Next

            End Using
        Catch ex As OleDbException
            MessageBox.Show("Error saving Testing to database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            ' MessageBox.Show("Error saving Testing to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("Unexpected Error: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub Expenseses_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        DataGridView1.Columns.Clear() ' Clear existing columns if any  

        Dim connection As New OleDbConnection(connectionString)
        ' Check database connectivity  
        Try
            ' Create a new OleDbConnection object and open the connection  

            connection.Open()

            ' Display the connection status on a button with a green background  
            Label8.Text = "Connected"
            Label8.BackColor = System.Drawing.Color.Green
            Label8.ForeColor = System.Drawing.Color.White
        Catch ex As Exception
            ' Display the connection status on a button with a red background  
            Label8.Text = "Not Connected"
            Label8.BackColor = System.Drawing.Color.Red
            Label8.ForeColor = System.Drawing.Color.White

            ' Display an error message  
            MessageBox.Show("Error connecting to the database" & ex.Message, "Database Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Close the database connection  
            connection.Close()
        End Try

        ' Disable certain buttons if the connection is not established  
        'Button3.Enabled = Label7.Text = "Connected"
        'Button4.Enabled = Label7.Text = "Connected"

        'LoadPersonDataFromDatabase()

    End Sub
End Class