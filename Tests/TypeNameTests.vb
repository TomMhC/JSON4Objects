

Namespace Data

    Public Class DeclaringClass

        Public Class ChildClass

        End Class

    End Class

    Public Class Generic(Of T)

    End Class

End Namespace

<TestClass()>
Public Class TypeNameTests

    <TestMethod()>
    Public Sub TestSimple()

        Dim t = GetType(Data.DeclaringClass)
        Dim typeName = JSON4Objects.Helpers.GetTypeNameFor(t)

        Dim t2 = Type.GetType(typeName)

        Assert.IsNotNull(t2)

        Assert.AreEqual(t.FullName, t2.FullName)

    End Sub

    <TestMethod()>
    Public Sub TestGeneric()

        Dim t = GetType(Data.Generic(Of Data.DeclaringClass))
        Dim typeName = JSON4Objects.Helpers.GetTypeNameFor(t)

        Dim t2 = Type.GetType(typeName)

        Assert.IsNotNull(t2)

        Assert.AreEqual(t.FullName, t2.FullName)

    End Sub

    <TestMethod()>
    Public Sub TestChild()

        Dim t = GetType(Data.DeclaringClass.ChildClass)
        Dim typeName = JSON4Objects.Helpers.GetTypeNameFor(t)

        Dim t2 = Type.GetType(typeName)

        Assert.IsNotNull(t2)

        Assert.AreEqual(t.FullName, t2.FullName)

    End Sub

End Class
