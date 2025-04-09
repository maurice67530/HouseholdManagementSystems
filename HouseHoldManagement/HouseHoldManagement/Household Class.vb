Public Class Household_Class

End Class
Public Class Register
    'properties of the Register class
    Public Property FullNames As String
    Public Property Password As Integer
    Public Property Username As String

    Public Property Email As String
    Public Property Role As String
    Public Property DateCreated As String
End Class

Public Class Logging
    ' Properties of the Login class
    Public Property Username As String
    Public Property Password As Integer

End Class

Public Class chores_
    Public Property Title As String
    Public Property AssignedTo As String
    Public Property Priority As String
    Public Property Status As String
    Public Property Frequency As String
    Public Property DueDate As String
    Public Property Recurring As String
    Public Property Description As String


End Class
Public Class MealPlans

    'properties of the mealplan class
    Public Property MealPlanId As String
    Public Property StartDate As DateTime
    Public Property EndDate As DateTime
    Public Property Meals As String
    Public Property MealName As String
    Public Property Items As String
    Public Property TotalCalories As Integer
    Public Property Description As String
    Public Property FilePath As String
    Public Property Calories As String

    Public Property Frequency As String

End Class
Public Class Inventory1

    Public Property ItemName As String
    Public Property Description As String
    Public Property Quantity As String
    Public Property Category As String
    Public Property ReoderLevel As String
    Public Property PricePerUnit As String
    Public Property DateAdded As DateTime
    Public Property ExpiryDate As DateTime
    Public Property Unit As String

End Class
Public Class Expensetracking

    Public Property ExpenseID As Integer
    Public Property Amount As Integer
    Public Property TotalIncome As Integer
    Public Property Description As String
    Public Property Tags As String
    Public Property Currency As String
    Public Property Category As String
    Public Property Paymentmethod As String
    Public Property Frequency As String
    Public Property ApprovalStatus As String
    Public Property Receiver As String
    Public Property DateOfexpenses As DateTime

End Class
Public Class DailyTask
    Public Property TaskID As String
    Public Property Title As String
    Public Property Description As String
    Public Property DueDate As DateTime
    Public Property Priority As String
    Public Property Status As String
    Public Property AssignedTo As String
End Class
