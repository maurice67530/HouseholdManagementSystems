Imports System.IO
Imports System.Windows.Forms

Module Module1
    Public personal As New List(Of person)
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\user\Documents\Visual Studio 2010\Projects\My hello\HOUSEHOLD MANAGEMENT SYSTEM.vb1\HOUSEHOLD MANAGEMENT SYSTEM.vb1\Austin.accdb"

    'Public Sub LoadPersonDataFromFile(filePath As String)
    '    If File.Exists(filePath) Then
    '        Try
    '            Using reader As New StreamReader(filePath)
    '                Dim person As person = Nothing
    '                While Not reader.EndOfStream
    '                    Dim line As String = reader.ReadLine()
    '                    If line.StartsWith("Name: ") Then
    '                        If person IsNot Nothing Then
    '                            personal.Add(person) ' Add the previous person to the list first  
    '                        End If
    '                        person = New person()
    '                        person.FirstName = line.Substring("Name: ".Length).Trim()
    '                    ElseIf line.StartsWith("Surname: ") Then
    '                        person.Lastname = line.Substring("Surname: ".Length).Trim()
    '                    ElseIf line.StartsWith("Contact Number: ") Then
    '                        person.ContactNumber = Integer.Parse(line.Substring("Contact Number: ".Length).Trim())
    '                    ElseIf line.StartsWith("eMail: ") Then
    '                        person.Email = line.Substring("eMail: ".Length).Trim()
    '                    ElseIf line.StartsWith("Role: ") Then
    '                        person.Role = line.Substring("Role: ".Length).Trim()
    '                    ElseIf line.StartsWith("Interest: ") Then
    '                        person.Interest = line.Substring("Interest: ".Length).Trim()
    '                    ElseIf line.StartsWith("Date Of Birth: ") Then
    '                        person.DateOfBirth = line.Substring("Date Of Birth: ".Length).Trim()
    '                    ElseIf line.StartsWith("Gender: ") Then
    '                        person.Gender = line.Substring("Gender: ".Length).Trim()
    '                    End If
    '                End While

    '                ' Add the last person to the list if we reached the end of the file  

    '                If person IsNot Nothing Then
    '                    personal.Add(person)
    '                End If
    '            End Using
    '        Catch ex As Exception
    '            Windows.Forms.MessageBox.Show("Error loading personnel data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        End Try
    '    Else
    '        MessageBox.Show("Personnel file not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End If
    ''    End 
    'End Sub
    Public Function GetPersonnelList() As List(Of person)
        Return personal
    End Function
    '    Private Sub LoadExpensesDataFromFile(filePath As String)
    '        If File.Exists(filePath) Then
    '            Try
    '                Using reader As New StreamReader(filePath)
    '                    Dim Expenses As Expenses = Nothing
    '                    While Not reader.EndOfStream
    '                        Dim line As String = reader.ReadLine()
    '                        If line.StartsWith("Name: ") Then
    '                            If Expenses IsNot Nothing Then
    '                                Expense.Add(Expenseses) ' Add the previous person to the list first  
    '                            End If
    '                            expense = New Expenses()
    '                            Expenses.Dates = line.Substring("Dates: ".Length).Trim()
    '                        ElseIf line.StartsWith("amount: ") Then
    '                            Expenses.Amount = line.Substring("amount: ".Length).Trim()
    '                        ElseIf line.StartsWith("Currency ") Then
    '                            Expenses.frequency = Integer.Parse(line.Substring("Frequency: ".Length).Trim())
    '                        ElseIf line.StartsWith("ApprovalStatus: ") Then
    '                            Expenses.approvalStatus = line.Substring("approvalstatus: ".Length).Trim()
    '                        ElseIf line.StartsWith("tags: ") Then
    '                            Expenses.Tags = line.Substring("Tags: ".Length).Trim()
    '                        ElseIf line.StartsWith("paymentmethod: ") Then
    '                            Expenses.payment = line.Substring("paymentmethod: ".Length).Trim()

    '                        End If
    '                    End While

    '                    Add the last person to the list if we reached the end of the file  

    '                    If person IsNot Nothing Then
    '                        expense.Add(Expenses)
    '                    End If
    '                End Using
    '            Catch ex As Exception
    '                MessageBox.Show("Error loading personnel data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            End Try
    '        Else
    '            MessageBox.Show("Personnel file not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        End If
End Module

