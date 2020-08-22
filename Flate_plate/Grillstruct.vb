Public Structure Grillstruct
    Public girder As Integer      'Girder no
    Public beam As Integer        'Beam no
    Public weight As Double       'Total weight

    Public Overrides Function Equals(obj As Object) As Boolean
        Throw New NotImplementedException()
    End Function

    Public Overrides Function GetHashCode() As Integer
        Throw New NotImplementedException()
    End Function

    Public Shared Operator =(left As Grillstruct, right As Grillstruct) As Boolean
        Return left.Equals(right)
    End Operator

    Public Shared Operator <>(left As Grillstruct, right As Grillstruct) As Boolean
        Return Not left = right
    End Operator
End Structure
