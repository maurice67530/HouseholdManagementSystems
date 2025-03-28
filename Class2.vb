Public Class Class2
End Class
Public Class person
        Public Property FirstName As String
        Public Property Lastname As String
        Public Property DateOfBirth As DateTime
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
        Public Property Quantity As Integer
        Public Property NutritionValue As Integer
        Public Property Category As String
        Public Property Price As Integer
        Public Property Unit As String
        Public Property ShelfLife As String
        Public Property ExpiryDate As DateTime
        Public Property ispurchased As String

    End Class
    '

    Public Class TaskManagement
        'properties of the dailytask class
        Public Property TaskId As Integer
        Public Property Title As String
        Public Property Description As String
        Public Property DueDate As DateTime
        Public Property priority As Integer
        Public Property Status As Integer
        Public Property AssignedTo As String
    End Class
    Public Class Expenses
        Friend priceperUnit As Object
        'properties of Expenses class

        Public Property ItemID As String
        Public Property Amount As String
        Public Property Currency As String
        Public Property Payment As String
        Public Property Tags As String
        Public Property Frequency As String
        Public Property ApprovalStatus As String
        Public Property Edit As DateTime
        Public Property ReorderLevel As String
        Public Property Delete As String
        Public Property Surname As String
    End Class
    Public Class Chores
        'properties of chores class 
        Public Property choresID As Integer
        Public Property Assigned As String
        Public Property Title As String
        Public Property Priority As String
        Public Property Status As String
        Public Property Frequency As String
        Public Property DueDate As DateTime
        Public Property Recurring As String

    End Class
    Public Class Budget
        'properties of Budget class
        Public Property paymentMethod As String
        Public Property PayingList As String
        Public Property Amount As Integer
        Public Property category As String
        Public Property Quantity As Integer
    End Class
    Public Class Inventory
        Public Property ItemID As Integer
        Public Property ItemName As String
        Public Property Description As String
        Public Property Quantity As Integer
        Public Property Unit As Integer
        Public Property Category As String
        Public Property ReorderLevel As Integer
        Public Property PricePerItem As Decimal
        Public Property DateAdded As DateTime
    End Class
    Public Class photogallery
        'properties of photogallery class
        Public Property DateoftheDay As DateTime
        Public Property paymentMethod As String
        Public Property priceperitem As Integer
        Public Property NumberOfPhototaken As Integer
        Public Property Description As String
        Public Property photoID As Integer
        Public Property Title As String
    Public Property filepath As String
    Public Property Tag As String
    Public Property Category As String
    Public Property Picture As String

End Class
    Public Class MealPlan
        'properties of Mealplan
        Public Property mealPlanID As Integer
        Public Property NameOfMeal As String
        Public Property StartDate As DateTime
    Public Property MealsType As String
    Public Property EndDate As DateTime
        Public Property totalcalories As Integer
        Public Property Description As String
        Public Property Picture As String
    'Public Property filter As String

End Class
    Public Class LOGIN
        'properties of LOGIN 
        Public Property FirstName As String
        Public Property LastName As String
        Public Property Role As String
    End Class
    Public Class regis
        'properties of REGISTER
        Public Property ID As String
        Public Property FirstName As String
        Public Property LastName As String
        Public Property Username As String
        Public Property password As String
        Public Property Email As String
        Public Property Role As String
        Public Property DateTime As DateTimePicker

        Public Class MealRecipes
            'properties of MealRecipes
            Public Property MealID As String
            Public Property MealName As String
            Public Property Ingredients As String
        End Class
    End Class




