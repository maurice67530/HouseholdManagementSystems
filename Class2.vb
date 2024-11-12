Public Class Class2

End Class

Public Class person
    Public Property FirstName As String
    Public Property Lastname As String
    Public Property DateOfBirth As String
    Public Property Gender As String
    Public Property ContactNumber As String
    Public Property Email As String
    Public Property Role As String
    Public Property MaritalStatus As String
    Public Property Interest As String


End Class
Public Class GroceryItem
    'properties of the GroceryITem class

    Public Property ItemName As String
    Public Property Quantity As String
    Public Property NutritionValue As String
    Public Property Category As String
    Public Property Price As String
    Public Property Unit As String
    Public Property ShelfLife As String
    Public Property ExpiryDate As String
    Public Property ispurchased As String

End Class
'

Public Class Dailytask
    'properties of the dailytask class
    Public Property TaskId As Integer
    Public Property Title As String
    Public Property Description As String
    Public Property priority As String
    Public Property status As String
    Public Property Assigned As String
End Class
Public Class Expenses
    'properties of Expenses class

    Public Property Amount As String
    Public Property Currency As String
    Public Property payment As String
    Public Property Tags As String
    Public Property frequency As String
    Public Property ApprovalStatus As String
    Public Property Dates As DateTime



End Class
Public Class Chores
    'properties of chores class 
    Public Property personNAme As String
    Public Property personLastName As String
    Public Property Role As String
    Public Property Dateofchores As Windows.Forms.DateTimePicker
    Public Property Typeofchores As String
End Class
Public Class Budget
    'properties of Budget class
    Public Property paymentMethod As String
    Public Property PayingList As String
    Public Property Amount As Decimal
    Public Property category As String
    Public Property Quantity As Integer
End Class
Public Class Inventory
    'properties of Inventory class 
    Public Property Typeofproducts As String
    Public Property AmountofEachproduct As Decimal
    Public Property productbelongto As String
    Public Property Quantityoftheproduct As Integer
    Public Property Categoryoftheproduct As String

End Class
Public Class photogallery
    'properties of photogallery class
    Public Property DateoftheDay As String
    Public Property paymentMethod As String
    Public Property priceperitem As Decimal
    Public Property numbersofPhototaken As Integer

End Class
Public Class MealPlan
    'properties of Mealplan 
    Public Property TypeOfMeal As String
    Public Property dateOfMeal As DateTimePicker
    Public Property amountOfmeal As ComboBox
    Public Property Unit As ComboBox
    Public Property mealPlanID As Integer
    Public Property Names As String
    Public Property StartDate As DateTime
    Public Property EndDate As DateTime
    Public Property Meals As List(Of MealPlan)
    Public Property totalcalories As Integer
    Public Property Description As String

End Class
Public Class LOGIN
    'properties of LOGIN 
    Public Property username As String
    Public Property password As String
End Class
Public Class REGISTER
    'properties of REGISTER
    Public Property IDNumber As String
    Public Property username As String
    Public Property Surname As String
    Public Property Dates As DateTimePicker
    Public Property Year As ComboBox

End Class
