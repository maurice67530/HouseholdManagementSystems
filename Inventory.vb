Imports System.Windows.Forms
Imports System.Data.OleDb
Imports HOUSEHOLD_MANAGEMENT_SYSTEM.vb

Public Class Inventory
    Private Expenses As IEnumerable(Of Expenses)

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DashBoard.Show()
        Me.Close()

    End Sub
    'Private Sub populatedDatagridview()

    '    Try
    '        Debug.WriteLine("populating datagridview successfully")
    '        ' Clear Existing rows   
    '        DataGridView1.Rows.Clear()

    '        ' Add each expense to the DataGridView  
    '        For Each exp As  In 


    '            DataGridView1.Rows.Add(exp.ItemID, exp.Names, exp.Description.ToString, exp.Quantity, exp.Unit, exp.Category, exp.Quantity, exp.ReorderLevel, exp.DateAdded())
    '        Next
    '    Catch ex As Exception
    '        Debug.WriteLine("fail to populate")
    '        Debug.WriteLine($"Stack Trace:{ex.StackTrace}")
    '        MessageBox.Show("An unexpexted error occured during DataGridview")

    '    End Try

    'End Sub

    Private Sub Inventory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadInventorydataFromDatabase()

        ' Set tooltips For buttons
        ToolTip1.SetToolTip(Button1, "Dashboard")
        ToolTip1.SetToolTip(ButtonSave, "Save")
        ToolTip1.SetToolTip(BtnEdit, "Edit")
        ToolTip1.SetToolTip(BtnDelete, "Delete")
        ToolTip1.SetToolTip(Button5, "Sort")
        ToolTip1.SetToolTip(Button6, "Filter")
        ToolTip1.SetToolTip(Button7, "Highlight")
        ToolTip1.SetToolTip(Button8, "Calculate")




    End Sub

    Public Sub loadInventorydataFromDatabase()
        Try
            Debug.WriteLine("populated connected succesfully")
            Using conn As New OleDbConnection(Module1.connectionString)
                conn.Open()

                Dim tableName As String = "Inventory"
                Dim cmd As New OleDbCommand($"SELECT * FROM {tableName}", conn)

                Dim da As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)

                DataGridView1.DataSource = dt

            End Using

        Catch ex As Exception

            Debug.WriteLine("populated fail to initialize succesfully")
            MessageBox.Show($"Error Loading  Expenses data from database: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ButtonSave_Click(sender As Object, e As EventArgs) Handles ButtonSave.Click
        BtnSaved()
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellMouseMove(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseMove

    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged

        'Load the data from  the selected row into UI controls 
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
            TextBox1.Text = selectedRow.Cells("ItemID").Value.ToString()
            TextBox2.Text = selectedRow.Cells("ItemName").Value.ToString()
            TextBox3.Text = selectedRow.Cells("Description").Value.ToString()
            TextBox4.Text = selectedRow.Cells("Quantity").Value.ToString()
            TextBox5.Text = selectedRow.Cells("Unit").Value.ToString()
            TextBox6.Text = selectedRow.Cells("Category").Value.ToString()
            TextBox7.Text = selectedRow.Cells("ReorderLevel").Value.ToString
            TextBox7.Text = selectedRow.Cells("PriceperItem").Value.ToString
            DateTimePicker1.Text = selectedRow.Cells("DateAdded").Value.ToString()
        End If

    End Sub


    Private Sub BtnEdit_Click_1(sender As Object, e As EventArgs) Handles BtnEdit.Click

        'BtnEdit()

    End Sub

    Private Sub BtnDelete_Click(sender As Object, e As EventArgs) Handles BtnDelete.Click

        'BtnDelete()

    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

    End Sub
End Class
