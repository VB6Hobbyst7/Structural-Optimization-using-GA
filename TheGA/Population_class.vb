Friend Class Population_class

    Private Val As Boolean


    Public Sub New()
        Val = False
    End Sub

    Public Property Value() As Byte

        Get
            If Val Then
                Return 1
            Else
                Return 0
            End If
        End Get

        'if input zero then false in all other cases it's true
        Set(value As Byte)
            If value = 0 Then
                Val = False
            Else
                Val = True
            End If
        End Set

    End Property

    Public Shared Operator +(ByVal Var As Population_class, ByVal Add As Integer) As Population_class

        Dim Add_operator As New Population_class


        If Var.Val Then
            If (Add Mod 2) = 0 Then
                Add_operator.Val = True
            Else
                Add_operator.Val = False
            End If
        Else
            If (Add Mod 2) = 0 Then
                Add_operator.Val = False
            Else
                Add_operator.Val = True
            End If
        End If


        Return Add_operator

    End Operator

    Public Shared Operator *(ByVal var As Population_class, ByVal mult As Double) As Double
        If var.Val Then
            Return mult
        Else
            Return 0 
        End If
    End Operator




End Class
