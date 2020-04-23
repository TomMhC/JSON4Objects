Imports System.Text

Namespace Data

    Public Class Person
        Public Property Name As String
        Public Property PreName As String
        Public Property Job As String
        Public Property Gender As Gender
        Public Property Age As Int32?
    End Class

    Public Enum Gender
        Male
        Female
    End Enum

End Namespace

<TestClass()>
Public Class Tests

#Region "Helpers"

    Public Function CompareEquality(obj1 As Hashtable, obj2 As Hashtable) As Boolean

        For Each k In obj1.Keys
            If Not obj2(k).ToString() = obj1(k).ToString() Then Return False
        Next
        Return True

    End Function

    Public Function CompareEquality(objectsCompared As IList(Of Object), obj1 As Object, obj2 As Object) As Boolean

        If objectsCompared.Contains(obj1) Then Return True
        objectsCompared.Add(obj1)

        If obj1 Is obj2 Then Return True
        If obj1 Is Nothing AndAlso Not obj2 Is Nothing Then Return False
        If obj2 Is Nothing AndAlso Not obj1 Is Nothing Then Return False
        If obj1 Is Nothing AndAlso obj2 Is Nothing Then Return True

        For Each p In obj1.GetType().GetProperties().Where(Function(prop) prop.GetIndexParameters().Count = 0)
            Dim p2 = obj2.GetType().GetProperty(p.Name, p.PropertyType)

            If p2 Is Nothing Then Return False

            Dim val1 = p.GetValue(obj1, Nothing)
            Dim val2 = p2.GetValue(obj2, Nothing)

            If val1 Is Nothing Then
                Return val2 Is Nothing
            ElseIf p.PropertyType.IsValueType Then
                If Not DirectCast(val1, IComparable).Equals(val2) Then Return False
            Else
                If Not CompareEquality(objectsCompared, val1, val2) Then Return False
            End If
        Next

        Return True

    End Function

#End Region

#Region "Simple Values"

    <TestMethod()>
    Public Sub TestString()

        Dim testVal = "Schaap\nSchaap"

        Using ms As New IO.MemoryStream
            Dim ser As New JSON4Objects.Serializer()
            ser.Serialize(ms, testVal)

            ms.Seek(0, IO.SeekOrigin.Begin)

            Using sr = New IO.StreamReader(ms)
                Assert.AreEqual(sr.ReadToEnd(), """Schaap\\nSchaap""")
            End Using
        End Using

    End Sub

    <TestMethod()>
    Public Sub TestInteger()

        Dim testVal = 24

        Using ms As New IO.MemoryStream
            Dim ser As New JSON4Objects.Serializer()
            ser.Serialize(ms, testVal)

            ms.Seek(0, IO.SeekOrigin.Begin)

            Using sr = New IO.StreamReader(ms)
                Assert.AreEqual(sr.ReadToEnd(), "24")
            End Using
        End Using

    End Sub

    <TestMethod()>
    Public Sub TestBoolean()

        Dim testVal = True

        Using ms As New IO.MemoryStream
            Dim ser As New JSON4Objects.Serializer()
            ser.Serialize(ms, testVal)

            ms.Seek(0, IO.SeekOrigin.Begin)

            Using sr = New IO.StreamReader(ms)
                Assert.AreEqual(sr.ReadToEnd(), "true")
            End Using
        End Using

    End Sub

    <TestMethod()>
    Public Sub TestDateTime()

        Dim testVal = New DateTime(2015, 2, 14, 23, 40, 15)

        Using ms As New IO.MemoryStream
            Dim ser As New JSON4Objects.Serializer()
            ser.Serialize(ms, testVal)

            ms.Seek(0, IO.SeekOrigin.Begin)

            Dim deSer As New JSON4Objects.Serializer()
            Dim obj2 = deSer.DeserializeDateTime(ms)

            Assert.AreEqual(testVal, obj2)
        End Using

    End Sub

#End Region

#Region "Constructions"

    <TestMethod()>
    Public Sub TestArray()

        Dim testVal = {4, 8, 15, 16, 23, 42}

        Using ms As New IO.MemoryStream
            Dim ser As New JSON4Objects.Serializer()
            ser.Serialize(ms, testVal)

            ms.Seek(0, IO.SeekOrigin.Begin)

            Using sr = New IO.StreamReader(ms)
                Assert.AreEqual(sr.ReadToEnd().Replace(" ", "").Replace(Environment.NewLine, ""), "[4,8,15,16,23,42]")
            End Using
        End Using

    End Sub

    <TestMethod()>
    Public Sub TestSimpleTypeObject()

        Using ms As New IO.MemoryStream
            Dim obj1 = New Data.Person With {.Name = "Caesar", .PreName = "Julius", .Job = "Emperor"}

            Dim ser As New JSON4Objects.Serializer()
            ser.AdvancedSerialization = False ' Make sure we don't get confused with Id or Type information
            ser.Serialize(ms, obj1)

            ms.Seek(0, IO.SeekOrigin.Begin)

            Using sr = New IO.StreamReader(ms)
                Assert.AreEqual(sr.ReadToEnd().Replace(" ", "").Replace(Environment.NewLine, ""), "{""PreName"":""Julius"",""Job"":""Emperor"",""Name"":""Caesar"",""Age"":null,""Gender"":0}")
            End Using
        End Using

    End Sub

    <TestMethod()>
    Public Sub TestSimpleTypeObjectComparisson()

        Using ms As New IO.MemoryStream
            Dim obj1 = New Data.Person With {.Name = "Caesar", .PreName = "Julius", .Job = "Emperor"}

            Dim ser As New JSON4Objects.Serializer()
            ser.Serialize(ms, obj1)

            ms.Seek(0, IO.SeekOrigin.Begin)

            Dim deSer As New JSON4Objects.Serializer()
            Dim obj2 = deSer.Deserialize(Of Data.Person)(ms)

            Assert.IsTrue(CompareEquality(New List(Of Object), obj1, obj2))
        End Using

    End Sub

    <TestMethod()>
    Public Sub TestSimpleTypeObjectWithNullable()

        Using ms As New IO.MemoryStream
            Dim obj1 = New Data.Person With {.Name = "Caesar", .PreName = "Julius", .Job = "Emperor", .Age = 30}

            Dim ser As New JSON4Objects.Serializer()
            ser.AdvancedSerialization = False ' Make sure we don't get confused with Id or Type information
            ser.Serialize(ms, obj1)

            ms.Seek(0, IO.SeekOrigin.Begin)

            Using sr = New IO.StreamReader(ms)
                Assert.AreEqual(sr.ReadToEnd().Replace(" ", "").Replace(Environment.NewLine, ""), "{""PreName"":""Julius"",""Job"":""Emperor"",""Name"":""Caesar"",""Age"":30,""Gender"":0}")
            End Using
        End Using

    End Sub

    <TestMethod()>
    Public Sub TestAnonymousTypes()

        Using ms As New IO.MemoryStream
            Dim obj1 = New With {.Name = "Caesar", .PreName = "Julius", .Job = "Emperor"}

            Dim ser As New JSON4Objects.Serializer()
            ser.Serialize(ms, obj1)

            ms.Seek(0, IO.SeekOrigin.Begin)

            Dim deSer As New JSON4Objects.Serializer()
            Dim obj2 = deSer.Deserialize(ms)

            Assert.IsTrue(CompareEquality(New List(Of Object), obj1, obj2))
        End Using

    End Sub

    <TestMethod()>
    Public Sub TestDeepAnonymousType()

        Using ms As New IO.MemoryStream
            Dim objJob = New With {.Description = "Emperor", .People = New List(Of Object)}
            Dim obj1 = New With {.Name = "Caesar", .PreName = "Julius", .Job = objJob}
            objJob.People.Add(obj1)

            Dim ser As New JSON4Objects.Serializer()
            ser.Serialize(ms, obj1)

            ms.Seek(0, IO.SeekOrigin.Begin)

            Dim deSer As New JSON4Objects.Serializer()
            Dim obj2 = deSer.Deserialize(ms)

            Assert.IsTrue(CompareEquality(New List(Of Object), obj1, obj2))
        End Using

    End Sub

    <TestMethod()>
    Public Sub TestObjectAsHashTable()

        Using ms As New IO.MemoryStream
            Dim obj1 = New With {.Name = "Caesar", .PreName = "Julius", .Job = "Emperor"}

            Dim ser As New JSON4Objects.Serializer()
            ser.AdvancedSerialization = False
            ser.Serialize(ms, obj1)

            ms.Seek(0, IO.SeekOrigin.Begin)

            Dim deSer As New JSON4Objects.Serializer()
            Dim obj2 = deSer.Deserialize(ms)

            Dim comparisonHt = New Hashtable(3)
            comparisonHt.Add("Name", "Caesar")
            comparisonHt.Add("PreName", "Julius")
            comparisonHt.Add("Job", "Emperor")

            Assert.IsTrue(CompareEquality(comparisonHt, CType(obj2, Hashtable)))

        End Using

    End Sub

#End Region

#Region "Custom Serializers"

    <TestMethod()>
    Public Sub TestList()

        Using ms As New IO.MemoryStream
            Dim obj1 = New With {.FamousBattles = New List(Of String)}
            obj1.FamousBattles.Add("Battle of Ravenna")
            obj1.FamousBattles.Add("Battle of Lacus Benacus")
            obj1.FamousBattles.Add("Battle of Carthage")

            Dim ser As New JSON4Objects.Serializer()
            ser.Serialize(ms, obj1)

            ms.Seek(0, IO.SeekOrigin.Begin)

            Dim deSer As New JSON4Objects.Serializer()
            Dim obj2 = deSer.Deserialize(ms)

            Assert.IsTrue(CompareEquality(New List(Of Object), obj1, obj2))

        End Using

    End Sub

    <TestMethod()>
    Public Sub TestDictionary()

        Using ms As New IO.MemoryStream
            Dim obj1 = New With {.FamousBattles = New Dictionary(Of String, String)}
            obj1.FamousBattles.Add("Battle of Ravenna", "Bonifatius")
            obj1.FamousBattles.Add("Battle of Lacus Benacus", "Claudius II")
            obj1.FamousBattles.Add("Battle of Carthage", "Maximinus ")

            Dim ser As New JSON4Objects.Serializer()
            ser.Serialize(ms, obj1)

            ms.Seek(0, IO.SeekOrigin.Begin)

            Dim deSer As New JSON4Objects.Serializer()
            Dim obj2 = deSer.Deserialize(ms)

            Assert.IsTrue(CompareEquality(New List(Of Object), obj1, obj2))

        End Using

    End Sub

    <TestMethod()>
    Public Sub TestListOfValueObject()

        Using ms As New IO.MemoryStream
            Dim obj1 = New With {.Numbers = New List(Of Integer)}
            obj1.Numbers.Add(1)
            obj1.Numbers.Add(2)
            obj1.Numbers.Add(3)

            Dim ser As New JSON4Objects.Serializer()
            ser.Serialize(ms, obj1)

            ms.Seek(0, IO.SeekOrigin.Begin)

            Dim deSer As New JSON4Objects.Serializer()
            Dim obj2 = deSer.Deserialize(ms)

            Assert.IsTrue(CompareEquality(New List(Of Object), obj1, obj2))

        End Using

    End Sub

    <TestMethod()>
    Public Sub TestListOfEnum()

        Using ms As New IO.MemoryStream
            Dim obj1 = New With {.Numbers = New List(Of Data.Gender)}
            obj1.Numbers.Add(Data.Gender.Male)
            obj1.Numbers.Add(Data.Gender.Female)
            obj1.Numbers.Add(Data.Gender.Male)

            Dim ser As New JSON4Objects.Serializer()
            ser.Serialize(ms, obj1)

            ms.Seek(0, IO.SeekOrigin.Begin)

            Dim deSer As New JSON4Objects.Serializer()
            Dim obj2 = deSer.Deserialize(ms)

            Assert.IsTrue(CompareEquality(New List(Of Object), obj1, obj2))

        End Using

    End Sub

    <TestMethod()>
    Public Sub TestDictOfValueObject()

        Using ms As New IO.MemoryStream
            Dim obj1 = New With {.Numbers = New Dictionary(Of Integer, Double)}
            obj1.Numbers.Add(1, 2)
            obj1.Numbers.Add(2, 4)
            obj1.Numbers.Add(3, 9)

            Dim ser As New JSON4Objects.Serializer()
            ser.Serialize(ms, obj1)

            ms.Seek(0, IO.SeekOrigin.Begin)

            Dim deSer As New JSON4Objects.Serializer()
            Dim obj2 = deSer.Deserialize(ms)

            Assert.IsTrue(CompareEquality(New List(Of Object), obj1, obj2))

        End Using

    End Sub

#End Region

    <TestMethod()>
    Public Sub TestDifferingTypes()

        Using ms As New IO.MemoryStream
            Dim monkey = New Animal With {.Species = Species.Mammal}

            Dim ser As New JSON4Objects.Serializer()
            ser.AdvancedSerialization = False
            ser.SetterHandling = JSON4Objects.Serializer.SetterHandlingEnum.FieldAlways
            ser.Serialize(ms, monkey)

            ms.Seek(0, IO.SeekOrigin.Begin)

            Dim deSer As New JSON4Objects.Serializer()
            Dim monkey2 = deSer.Deserialize(Of Animal)(ms)

            Assert.AreEqual(monkey.Species, monkey2.Species)

        End Using

    End Sub

    Public Class Animal

        Private _species As Species?
        Public Property Species As Species
            Get
                Return _species.GetValueOrDefault(Species.Fish)
            End Get
            Set(value As Species)
                _species = value
            End Set
        End Property

    End Class

    Public Enum Species
        Fish
        Mammal
        Bird
    End Enum

End Class
