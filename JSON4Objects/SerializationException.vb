Public Class SerializationException
    Inherits Exception

    Public Sub New(message As String, ex As Exception)
        MyBase.New(message, ex)
    End Sub

    Public Sub New(message As String, propertyName As String, ex As Exception)
        Me.New(message, ex)

        Me.PropertyName = propertyName
    End Sub

    Public Property PropertyName As String

End Class
