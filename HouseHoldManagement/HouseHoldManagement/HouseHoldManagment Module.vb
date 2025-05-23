﻿
Imports System.IO
Imports System.Data.OleDb
Module HouseHoldManagment_Module

    Public currentUser As String ' Global variable for logged-in user
    Public Property conn As New OleDbConnection(connectionString)
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Khodani\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb;Persist Security Info=False;"

    Public Function GetConnection() As OleDbConnection
        Return New OleDbConnection(connectionString)
    End Function

End Module

Module Ndamu
    Public Const connectionstring As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Nedzamba\Source\Repos\HouseholdManagementSystems\HMS.accdb"
    Public Sub AddEvent(Title As String, Notes As String, EventDate As Date, StartTime As Date, Endtime As Date, AssignedTo As String, EventType As String)
        Try
            Using conn As New OleDbConnection(Ndamu.connectionstring)
                conn.Open()
                Dim query As String = "INSERT INTO FamilySchedule ([Title],[ Notes], [DateOfEvent], [StartTime], [EndTime], [AssignedTo], [EventType])" & "VALUES (?, ?, ?, ?, ?, ?, ?)"
                Using cmd As New OleDbCommand(query, conn)
                    cmd.Parameters.AddWithValue("?", Title)
                    cmd.Parameters.AddWithValue("?", Notes)
                    cmd.Parameters.AddWithValue("?", EventDate)
                    cmd.Parameters.AddWithValue("?", StartTime)
                    cmd.Parameters.AddWithValue("?", Endtime)
                    cmd.Parameters.AddWithValue("?", AssignedTo)
                    cmd.Parameters.AddWithValue("?", EventType)
                    cmd.ExecuteNonQuery()

                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error adding schedule: " & ex.Message)
        End Try
    End Sub
    Public Sub AddChoreEvent(Title As String, AssignedTo As String, DueDate As Date)
        'Add Chore Event
        AddEvent(Title, "Chores automatically added", DueDate, DueDate, DueDate.AddHours(1), AssignedTo, "Chores")
    End Sub
    Public Sub AddMealEvent(MealName As String, StartDate As Date)
        'Add Meal Event
        AddEvent(MealName, "MealPlans", StartDate, StartDate.AddHours(12), StartDate.AddHours(13), "all", "Meal")
    End Sub
    Public Sub AddTaskReminder(Title As String, AssignedTo As String, DueDate As Date)
        'Add Task Reminder
        AddEvent(Title, "Task Due Reminder", DueDate, DueDate, DueDate.AddHours(1), AssignedTo, "Task")
    End Sub
    Public Sub MarkPhotoDay(DateAdded As Date, Description As String)
        'Add Photo Day Event
        AddEvent("Photo Day", Description, DateAdded, DateAdded.AddHours(15), DateAdded.AddHours(16), "Family", "Photo")
    End Sub
End Module

Module Rasta
    ' Connection string using relative path to the database
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Raphalalani\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"

    ' Function to create and return a connection object
    Public Function GetConnection() As OleDbConnection
        Return New OleDbConnection(connectionString)
    End Function
End Module
Module Module1
    Public Property conn As New OleDbConnection(connectionString)
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Zwivhuya\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"



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

End Module

Module Rinae
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Rinae\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"

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
End Module

'[Dongola] vhasongo silinga
Module Cruwza
    Public Property conn As New OleDbConnection(connectionString)
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Dongola\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"

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

End Module
Module Xiluva
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Xiluva\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"
End Module
'Murangi (don't Touch)
Module Murangi
    Public Const connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Murangi\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"
End Module
Module Rotondwa
    Public Property conn As New OleDbConnection(connectionString)
    Public Const connectionString As String = " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Rotondwa\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb;Persist Security Info=False;"

End Module
Module Ndivhuwo
    Public Property conn As New OleDbConnection(connectionString)
    Public Const connectionString As String = " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Delicious\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"

End Module
Module Masindi
    Public Property conn As New OleDbConnection(connectionString)
    Public Const connectionString As String = " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Masindi\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb;Persist Security Info=False;"
End Module
Module khodani
    Public Property conn As New OleDbConnection(connectionString)
    Public Const connectionString As String = " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\khodani\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"
End Module
Module Faith
    Public Property conn As New OleDbConnection(connectionString)
    Public Const connectionString As String = " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Faith\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"
End Module
Module Mulalo
    Public Property conn As New OleDbConnection(connectionString)
    Public Const connectionString As String = " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Austin\Source\Repos\maurice67530\HouseholdManagementSystems\HMS.accdb"

End Module


