Public Class Household_Class

End Class
Public Class Registration

    'properties of the Register class
    Public Property FullNames As String
    Public Property Password As Integer
    Public Property Username As String

    Public Property Email As String
    Public Property Role As String
    Public Property DateCreated As String
    Public Property Age As String
    Public Property Picture As String

End Class

Public Class Logging
    ' Properties of the Login class
    Public Property username As String
    Public Property Password As Integer

End Class

Public Class FamilyMembers
    Public Property NameOfPerson As String
    Public Property DoesNotEat As String
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
    Public Property EndTime As String
    Public Property StartTime As String

End Class
Public Class MealPlans
    'property of mealplan class
    Public Property Meals As String
    Public Property Items As String
    Public Property MealName As String
    Public Property StartDate As DateTime
    Public Property EndDate As DateTime
    Public Property TotalCalories As String
    Public Property Description As String
    Public Property Calories As String
    Public Property FilePath As String
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
    Public Property Person As String
    Public Property DateOfexpenses As String
    Public Property BillName As String
    Public Property StartDate As String
    Public Property Recurring As String
    Public Property paid As String
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
Public Class Groceryy
    'Properties of the GroceryIteam class
    Public Property ItemName As String
    Public Property Quantity As String
    Public Property Category As String
    Public Property Unit As String
    Public Property ExpiryDate As String
    Public Property PricePerUnit As String
    Public Property PurchaseDate As String
    Public Property Period As String
End Class
Public Class Person
    Public Property FirstName As String
    Public Property LastName As String
    Public Property Gender As String
    Public Property Email As String
    Public Property DateOfBirth As DateTime
    Public Property Role As String
    Public Property MaritalStatus As String
    Public Property postalcode As String
    Public Property Age As Integer
    Public Property Contact As Integer
    Public Property Photo As String
    Public Property Dietary As String
    Public Property FoodType As String


End Class
Public Class FamilySchedule
    Public Property ScheduleID As String
    Public Property Title As String
    Public Property Notes As String
    Public Property DateOfEvent As String
    Public Property StartTime As String
    Public Property EndTime As String
    Public Property AssignedTo As String
    Public Property EventType As String
    Public Property IsBudgetRequired As String
    Public Property Amount As Integer

End Class
Public Class BudgetClass
    Public Property Person As String
    Public Property Role As String
    'Public Property EndDate As String
    Public Property Income As String
    Public Property Utilities As String
    Public Property Rent As String
    Public Property Expenses As String
    Public Property Frequency As String
    Public Property EndDate As String
    Public Property Groceries As String
    Public Property StartDate As String
    Public Property Remaining As String
    Public Property BudgetAmount As String
    Public Property Amount As String

End Class
Public Class HouseholdDocument
    Public Property HouseholdID As String
    Public Property Title As String
    Public Property Notes As String
    Public Property Category As String
    Public Property FilePath As String
    Public Property UploadedBy As String
    Public Property UploadDate As String

End Class