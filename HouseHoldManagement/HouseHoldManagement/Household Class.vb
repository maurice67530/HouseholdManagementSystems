Public Class Household_Class
    Public Class expenses

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
End Class
Public Class Register
    'properties of the Register class
    Public Property FullNames As String
    Public Property Password As Integer
    Public Property Username As String
    Public Property ID As Integer
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