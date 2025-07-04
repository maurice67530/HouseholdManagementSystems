﻿
Imports System.IO
Imports System.Data.OleDb
Module HouseHoldManagment_Module


    Public currentUser As String ' Global variable for logged-in user
    Public Property conn As New OleDbConnection(connectionString)
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\MUDAUMURANGI\Users\Murangi\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"
    Public Sub InsertEvent(conn As OleDbConnection, title As String, notes As String, dateOfEvent As Date, startTime As Date, endTime As Date, assignedTo As String, eventType As String)
        Using insertCmd As New OleDbCommand("INSERT INTO FamilySchedule (Title, Notes, DateOfEvent, StartTime, EndTime, AssignedTo, EventType) VALUES (?, ?, ?, ?, ?, ?, ?)", conn)
            insertCmd.Parameters.AddWithValue("?", title)
            insertCmd.Parameters.AddWithValue("?", notes)
            insertCmd.Parameters.AddWithValue("?", dateOfEvent)
            insertCmd.Parameters.AddWithValue("?", startTime)
            insertCmd.Parameters.AddWithValue("?", endTime)
            insertCmd.Parameters.AddWithValue("?", assignedTo)
            insertCmd.Parameters.AddWithValue("?", eventType)
            insertCmd.ExecuteNonQuery()
        End Using
    End Sub
    Public Function EventExists(conn As OleDbConnection, title As String, dateOfEvent As Date) As Boolean
        Using checkCmd As New OleDbCommand("SELECT COUNT(*) FROM FamilySchedule WHERE Title = ? AND DateOfEvent = ?", conn)
            checkCmd.Parameters.AddWithValue("?", title)
            checkCmd.Parameters.AddWithValue("?", dateOfEvent)
            Return Convert.ToInt32(checkCmd.ExecuteScalar()) > 0
        End Using
    End Function
    Public Function GetSchoolTripDates() As List(Of Date)
        Dim results As New List(Of Date)
        Try
            Using conn As New OleDbConnection(connectionString)
                conn.Open()
                Dim cmd As New OleDbCommand("SELECT StartDate FROM Budget", conn)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        If Not IsDBNull(reader("StartDate")) Then
                            results.Add(CDate(reader("StartDate")))
                        End If
                    End While
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading School Trip dates: " & ex.Message)
        End Try
        Return results
    End Function

    Public Function GetDoctorVisitDates() As Dictionary(Of Date, String)
        Dim results As New Dictionary(Of Date, String)
        Try
            Using conn As New OleDbConnection(connectionString)
                conn.Open()
                Dim cmd As New OleDbCommand("SELECT DateOfexpenses, Person FROM Expense", conn)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        If Not IsDBNull(reader("DateOfexpenses")) AndAlso Not IsDBNull(reader("Person")) Then
                            Dim dt As Date = CDate(reader("DateOfexpenses"))
                            Dim person As String = reader("Person").ToString()
                            If Not results.ContainsKey(dt) Then
                                results.Add(dt, person)
                            End If
                        End If
                    End While
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading Doctor Visit dates: " & ex.Message)
            ' MessageBox.Show("Error saving Schedule to database: " & ex.Message & vbNewLine & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return results
    End Function
    Public Function Getconnection() As OleDbConnection
        Return New OleDbConnection(connectionString)
    End Function
    Public Sub FilterBudget(Frequency As String)
        Dim BudgetTable As New DataTable()


        Try
            conn.Open()
            Dim query As String = "Select * FROM Budget WHERE 1=1"

            If Not String.IsNullOrEmpty(Frequency) Then
                query &= " AND Frequency = @Frequency"
            End If

            'If Not String.IsNullOrEmpty(StartDate) Then
            '    query &= " AND StartDate = @StartDate"
            'End If

            Dim command As New OleDb.OleDbCommand(query, conn)

            If Not String.IsNullOrEmpty(Frequency) Then
                command.Parameters.AddWithValue("@Frequency", Frequency)
            End If

            'If Not String.IsNullOrEmpty(Calories) Then
            '    command.Parameters.AddWithValue("@StartDate", StartDate)
            'End If

            Dim adapter As New OleDb.OleDbDataAdapter(command)
            adapter.Fill(BudgetTable)
            Budget.DataGridView1.DataSource = BudgetTable

        Catch ex As Exception
            MsgBox("Error filtering Budget: " & ex.Message, MsgBoxStyle.Critical, "Database Error")
        Finally
            conn.Close()
        End Try
    End Sub

    Public Sub FilterMealPlan(Calories As String)
        Dim Mealtable As New DataTable()


        Try
            conn.Open()
            Dim query As String = "Select * FROM MealPlans WHERE 1=1"

            If Not String.IsNullOrEmpty(Calories) Then
                query &= " AND Calories = @Calories"
            End If

            'If Not String.IsNullOrEmpty(StartDate) Then
            '    query &= " AND StartDate = @StartDate"
            'End If

            Dim command As New OleDb.OleDbCommand(query, conn)

            If Not String.IsNullOrEmpty(Calories) Then
                command.Parameters.AddWithValue("@Calories", Calories)
            End If

            'If Not String.IsNullOrEmpty(Calories) Then
            '    command.Parameters.AddWithValue("@StartDate", StartDate)
            'End If

            Dim adapter As New OleDb.OleDbDataAdapter(command)
            adapter.Fill(Mealtable)
            MealPlan.DataGridView1.DataSource = Mealtable

        Catch ex As Exception
            MsgBox("Error filtering tasks: " & ex.Message, MsgBoxStyle.Critical, "Database Error")
        Finally
            conn.Close()
        End Try
    End Sub


    Public Sub ClearControls(ByVal form As Form)

        'clear TextBoxes
        For Each ctrl As Control In form.Controls
            If TypeOf ctrl Is TextBox Then
                CType(ctrl, TextBox).Clear()
            End If


        Next

        'clear comboBoxes
        For Each ctrl As Control In form.Controls
            If TypeOf ctrl Is ComboBox Then
                CType(ctrl, ComboBox).ResetText()
            End If
        Next

        'clear DateTimePicker
        For Each ctrl As Control In form.Controls
            If TypeOf ctrl Is DateTimePicker Then
                CType(ctrl, DateTimePicker).Value = DateTimePicker.MinimumDateTime
            End If
        Next

    End Sub

    Public Sub FilterInventory(Category As String, Unit As String)
        Dim taskTable As New DataTable
        Dim conn As New OleDbConnection(connectionString)
        Try
            conn.Open()
            Dim query As String = "SELECT * FROM Inventory WHERE 1=1"

            If Not String.IsNullOrEmpty(Category) Then
                query &= " AND Category = @Category"
            End If

            If Not String.IsNullOrEmpty(Unit) Then
                query &= " AND Unit = @Unit"
            End If

            Dim command As New OleDb.OleDbCommand(query, conn)

            If Not String.IsNullOrEmpty(Category) Then
                command.Parameters.AddWithValue("@Category", Category)
            End If

            If Not String.IsNullOrEmpty(Unit) Then
                command.Parameters.AddWithValue("@Unit", Unit)
            End If

            Dim adapter As New OleDb.OleDbDataAdapter(command)
            adapter.Fill(taskTable)
            Inventory.DataGridView1.DataSource = taskTable
        Catch ex As Exception
            MsgBox("Error filtering Inventory: " & ex.Message, MsgBoxStyle.Critical, "Database Error")

        Finally
            conn.Close()
        End Try
    End Sub
    Public Sub FilterPhoto(FamilyMember As String) ', DateAdded As DateTime)
        Dim taskTable As New DataTable
        Dim conn As New OleDbConnection(connectionString)
        Try
            conn.Open()
            Dim query As String = "SELECT * FROM Photos WHERE 1=1"

            If Not String.IsNullOrEmpty(FamilyMember) Then
                query &= " AND FamilyMember = @FamilyMember"
            End If

            'If Not String.IsNullOrEmpty(DateAdded) Then
            '    query &= " AND DateAdded = @DateAdded"
            'End If

            Dim command As New OleDb.OleDbCommand(query, conn)

            If Not String.IsNullOrEmpty(FamilyMember) Then
                command.Parameters.AddWithValue("@FamilyMember", FamilyMember)
            End If

            'If Not String.IsNullOrEmpty(DateAdded) Then
            '    command.Parameters.AddWithValue("@DateAdded", DateAdded)
            'End If

            Dim adapter As New OleDb.OleDbDataAdapter(command)
            adapter.Fill(taskTable)
            PhotoGallery.DataGridView1.DataSource = taskTable
        Catch ex As Exception
            MsgBox("Error filtering photo: " & ex.Message, MsgBoxStyle.Critical, "Database Error")
        Finally
            conn.Close()
        End Try
    End Sub

    Public Sub FilterChores(Frequency As String, Priority As String)
        Dim taskTable As New DataTable()
        Try
            Dim Query As String = "SELECT * FROM Chores WHERE 1 = 1"

            If Not String.IsNullOrEmpty(Frequency) Then
                Query &= " AND Frequency = @Frequency"
            End If
            If Not String.IsNullOrEmpty(Priority) Then
                Query &= " AND Priority = @Priority"
            End If

            Dim Command As New OleDb.OleDbCommand(Query, conn)

            If Not String.IsNullOrEmpty(Frequency) Then
                Command.Parameters.AddWithValue("@Frequency", Frequency)
            End If


            If Not String.IsNullOrEmpty(Priority) Then
                Command.Parameters.AddWithValue("@Priority", Priority)
            End If

            Dim Adapter As New OleDb.OleDbDataAdapter(Command)
            Adapter.Fill(taskTable)
            chores.DGVChores.DataSource = taskTable

        Catch ex As Exception
            MsgBox("Error Filtering tasks: " & ex.Message, MsgBoxStyle.Critical, "Database error")
        Finally
            conn.Close()
        End Try
    End Sub
    Public Sub CheckDatabaseConnection(statusLabel As Label)
        ' Access .accdb file on a network path
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\MUDAUMURANGI\Users\Murangi\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb;Persist Security Info=False;"

        Using conn As New OleDbConnection(connectionString)
            Try
                conn.Open()
                statusLabel.Text = "Database Connected"
                statusLabel.ForeColor = Color.Green
                statusLabel.BackColor = Color.BurlyWood
            Catch ex As Exception
                statusLabel.Text = "Database NOT Connected: " & ex.Message
                statusLabel.ForeColor = Color.Red
                statusLabel.BackColor = Color.BurlyWood
            End Try
        End Using
    End Sub
End Module


