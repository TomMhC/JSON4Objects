<AttributeUsage(AttributeTargets.Property Or AttributeTargets.Field, AllowMultiple:=False, Inherited:=True)> _
Public Class Ignore
    Inherits Attribute

    Public Sub New()
    End Sub
End Class

<AttributeUsage(AttributeTargets.Property Or AttributeTargets.Field, AllowMultiple:=True, Inherited:=True)> _
Public Class Pseudonym
    Inherits Attribute

    Public Property Pseudonym As String

    Public Sub New(ByVal pseudonym As String)
        Me.Pseudonym = pseudonym
    End Sub
End Class

<AttributeUsage(AttributeTargets.Property Or AttributeTargets.Field, AllowMultiple:=False, Inherited:=True)>
Public Class Setter
    Inherits Attribute

    Public Sub New(tSetterHandling As JSON4Objects.Serializer.SetterHandlingEnum)
        Me.SetterHandling = tSetterHandling
    End Sub

    Public Property SetterHandling As JSON4Objects.Serializer.SetterHandlingEnum

End Class