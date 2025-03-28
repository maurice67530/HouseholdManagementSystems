
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Windows.Forms.DataVisualization.Charting

Public Class DashBoard
    ' Create a ToolTip object
    Private toolTip As New ToolTip()

    Public Property TASKBAR1 As Object

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        person.ShowDialog()

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        GroceryItemvb.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click

        TaskManagement.ShowDialog()

    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click

        Expenses.ShowDialog()

    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click

        Chores.ShowDialog()
    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click

        Budjet.ShowDialog()
    End Sub

    Private Sub Button7_Click(sender As System.Object, e As System.EventArgs) Handles Button7.Click

        Inventory.ShowDialog()

    End Sub

    Private Sub Button8_Click(sender As System.Object, e As System.EventArgs) Handles Button8.Click

        PHOTO_GALLERY.ShowDialog()

    End Sub

    Private Sub Button9_Click(sender As System.Object, e As System.EventArgs) Handles Button9.Click
        MealPlan.ShowDialog()

    End Sub

    Private Sub Button10_Click(sender As System.Object, e As System.EventArgs)

        MY_USER_VALIDATION.ShowDialog()
    End Sub

    Private Sub Button11_Click(sender As System.Object, e As System.EventArgs) Handles Button11.Click
        'regis.ShowDialog()
    End Sub

    Private Sub Button12_Click(sender As System.Object, e As System.EventArgs)

    End Sub
    Private Sub Loadchartdata()
        'Update this connection string based on my database configuration

        Dim connectionString As String = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Austin\Documents\Visual Studio 2010\Projects\My hello\HOUSEHOLD MANAGEMENT SYSTEM.vb1\HOUSEHOLD MANAGEMENT SYSTEM.vb1\Austin.accdb"
        Dim query As String = "SELECT [status],[priority] FROM [TaskManagement]"

        Using conn As New OleDbConnection(connectionString)
            Dim command As New OleDbCommand(query, conn)
            conn.Open()

            Using reader As OleDbDataReader = command.ExecuteReader
                While reader.Read
                    'assuming column X is a string (category) and columnX is numeric value
                    Dim Status As String = reader("Status").ToString
                    Dim Priority As String = reader("priority").ToString

                    'add points to the charts;change the series name added
                    Chart1.Series("Series1").Points.AddXY(Status, Priority)

                    Chart1.ChartAreas(0).AxisX.Title = "your X axis Title"
                    Chart1.ChartAreas(0).AxisY.Title = "your Y axis Title"

                End While
            End Using
        End Using

        conn.Close()
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub DashBoard_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Loadchartdata()
        PopulateListboxFromdatabase()
        LOADCHORESSTATUS()
        UpdateBudgetStatus()
        LoadUpcomingMealPlans()
        PopulateTextbox2Fromdatabase()

    End Sub

    Private Sub PopulateListboxFromdatabase()


        Dim conn As New OleDbConnection(Module1.connectionString)
        Try
            Debug.WriteLine("Listbox populated from database")
            conn.Open()

            'Retrieve the status and title columns from the chore tabel
            Dim query As String = "SELECT Status,Title From CHORES "
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader

            'bind the retrieved data to the combobox 
            Listbox1.Items.Clear()
            While reader.Read()
                Listbox1.Items.Add($"{reader("Status")}{reader("Title")}")
            End While

            'Close the database
            reader.Close()

        Catch ex As Exception
            Debug.WriteLine("Failed to populate listbox")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            'handle any exeptions that may occur
            MessageBox.Show($"Error:{ex.Message}")

        Finally

            'close the database connection 
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
        Debug.WriteLine("Done with populating listbox from database ")
    End Sub
    Private Sub PopulateTextbox2Fromdatabase()


        Dim conn As New OleDbConnection(Module1.connectionString)
        Try
            Debug.WriteLine("Listbox populated from database")
            conn.Open()

            'Retrieve the status and title columns from the chore tabel
            Dim query As String = "SELECT Amount,Category From Budget "
            Dim cmd As New OleDbCommand(query, conn)
            Dim reader As OleDbDataReader = cmd.ExecuteReader

            'bind the retrieved data to the combobox 
            ListBox2.Items.Clear()
            While reader.Read()
                ListBox2.Items.Add($"{reader("Amount")}{reader("Category")}")
            End While

            'Close the database
            reader.Close()

        Catch ex As Exception
            Debug.WriteLine("Failed to populate listbox")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")
            'handle any exeptions that may occur
            MessageBox.Show($"Error:{ex.Message}")

        Finally

            'close the database connection 
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
        Debug.WriteLine("Done with populating listbox from database ")
    End Sub


    Private Sub LOADCHORESSTATUS()
        Dim completed As Integer = 0, Inprogress As Integer = 0, NotStarted As Integer = 1
        Dim Query As String = "SELECT status, COUNT(*) FROM CHORES group by status"
        Using conn As New OleDbConnection(connectionString), cmd As New OleDbCommand(Query, conn)

            conn.Open()

            Using reader = cmd.ExecuteReader
                While reader.Read
                    Select Case reader("Status").ToString
                        Case "Completed"
                            completed = Convert.ToUInt32(reader(1))
                        Case "Inprogress"
                        Case "NotStarted"
                            NotStarted = Convert.ToUInt32(reader(1))

                    End Select
                End While
            End Using
        End Using
        Label7.Text = $"chore--completed:{completed}| Inprogress:{Inprogress}|NotStarted:{NotStarted}"
    End Sub
    Private Sub DisplayPhoto()

        If photolist.Count > 0 Then
            FlowLayoutPanel1.Controls.Clear()
            Dim pb As New PictureBox()
            pb.Image = Image.FromFile(photolist(currentPhotoIndex))
            pb.SizeMode = PictureBoxSizeMode.StretchImage 'set stretch image
            pb.Size = FlowLayoutPanel1.Size 'Match panel size
            FlowLayoutPanel1.Controls.Add(pb) 'Add to flowLayoutPanel1
        End If
    End Sub
    Private Sub setupTimer()
        photoTimer.Interval = 2000 '2 seconds
        AddHandler photoTimer.Tick, AddressOf PhotoTimer_Tick
        photoTimer.Start()
    End Sub
    Private Sub PhotoTimer_Tick(sender As Object, e As EventArgs)
        If photolist.Count > 0 Then
            currentPhotoIndex = (currentPhotoIndex + 1) Mod photolist.Count
            DisplayPhoto()
        End If
    End Sub
    Private photolist As New List(Of String)() 'List to store photo path
    Private currentPhotoIndex As Integer = 0
    Private WithEvents photoTimer As New Timer()

    Private Sub LoadRecentPhotos()

        photolist.Clear()
        Dim query As String = "SELECT TOP 5 photo FROM photoAlbum ORDER BY Datetaken "
        Using conn As New OleDbConnection(connectionString)
            Using cmd As New OleDbCommand(query, conn)

                conn.Open()
                Using reader As OleDbDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        photolist.Add(reader("photos").ToString())
                    End While
                End Using
            End Using
        End Using

    End Sub
    Private Sub Chart1_Click(sender As Object, e As EventArgs) Handles Chart1.Click

    End Sub
    Private Sub LoadUpcomingMealPlans()

        Dim query As String = "SELECT COUNT(*)From MealPlan"
        Dim cmd As New OleDbCommand(query, conn)
        cmd.Parameters.AddWithValue("@Date", DateTime.Now)
        conn.Open()
        Dim upcomingMeals As Integer = CInt(cmd.ExecuteScalar())
        conn.Close()
        TextBox2.Text = "upcoming Meal Plans: " & upcomingMeals
    End Sub

    Private Sub UpdateBudgetStatus()

        Dim query As String = "SELECT SUM(Amount)FROM Expenses"
        Using Conn As New OleDbConnection(connectionString)
            Conn.Open()
            Dim cmd As New OleDbCommand(query, Conn)
            Dim totalExpenses As Decimal = Convert.ToDecimal(cmd.ExecuteScalar())
            'Assume you have a label for budget
            Label5.Text = "Expense budget: R" & totalExpenses.ToString()

            'Assuming a fixed budget,for example $500
            Dim Budget As Decimal = 2000
            Label6.Text = "Expense Budget In %: : " & ((totalExpenses \ Budget) * 100).ToString("F2") & "%"
            'Update a progressbar if you have one
            'TASKBAR1.Value = CInt((totalExpenses / Budget) * 100)
        End Using
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub FlowLayoutPanel1_Paint(sender As Object, e As PaintEventArgs) Handles FlowLayoutPanel1.Paint

    End Sub

    Private Sub Button10_Click_1(sender As Object, e As EventArgs) Handles Button10.Click
        EmailForm.ShowDialog()
    End Sub

    Private Sub Button12_Click_1(sender As Object, e As EventArgs) Handles Button12.Click
        InternetForm.ShowDialog()
    End Sub
End Class