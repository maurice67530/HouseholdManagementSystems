Public Class GroceryItemvb

    Dim Grocery_Items As vb.GroceryItem

    Private Sub btnDashboard_Click(sender As System.Object, e As System.EventArgs) Handles btnDashboard.Click
        Me.Close()
    End Sub
    Private Sub btnSubmit_Click(sender As System.Object, e As System.EventArgs) Handles btnSubmit.Click
        'create an instance of groceryItem
        Dim item As New GroceryItem

        'Assign values from Textboxes to the  GroceryItem properties 
        item.ItemName = TextBox1.Text
        item.Quantity = TextBox2.Text
        item.NutritionValue = TextBox3.Text
        item.Category = TextBox4.Text
        item.Price = TextBox5.Text
        item.Unit = TextBox6.Text
        item.ShelfLife = TextBox7.Text
        item.ExpiryDate = DateTimePicker1.Text
        item.ispurchased = TextBox8.Text

        'add the Grocery to the list

        'add the Grocery to the list
        groceryItem.Add(Grocery_Items)



        'Display a confirmation message 
        MsgBox("Grocery Item Added!" & vbCrLf &
        "ItemName:" & item.ItemName & vbCrLf &
        "Quantity:" & item.Quantity & vbCrLf &
        "NutriationValue:" & item.NutritionValue & vbCrLf &
        "Category:" & item.Category & vbCrLf &
        "Price:" & item.Price & vbCrLf &
        "Unit:" & item.Unit & vbCrLf &
        "ShelLife:" & item.ShelfLife & vbCrLf &
        "ExpiryDate:" & item.ExpiryDate & vbCrLf &
        "ispurchased:" & item.ispurchased, vbInformation, "Item confirmation")



        populatedDataGridview()
    End Sub
    Private groceryItem As New List(Of GroceryItem)
    Private Sub populatedDataGridview()
        'clear Existing rows
        DataGridView1.Rows.Clear()

        'Add each expense to the DataGridView
        For Each per As GroceryItem In groceryItem
            DataGridView1.Rows.Add(per.ItemName, per.Quantity, per.NutritionValue, per.Category, per.Price, per.Unit, per.ShelfLife.ToString, per.ExpiryDate, per.ispurchased)




        Next
    End Sub

    Private Sub GroceryItem_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        DataGridView1.Columns.Add("ItemName", "ItemName")
        DataGridView1.Columns.Add("Quantity", "Quantity")
        DataGridView1.Columns.Add("Nutritionvalue", "NutritionValue")
        DataGridView1.Columns.Add("Category", "Category")
        DataGridView1.Columns.Add("Price", "Price")
        DataGridView1.Columns.Add("Unit", "Unit")
        DataGridView1.Columns.Add("ShelfLife", "ShelfLife")
        DataGridView1.Columns.Add("ExpiryDate", "ExpiryDate")
        DataGridView1.Columns.Add("Ispurchased", "Ispurchased")

    End Sub
End Class



