Imports System.Threading.Tasks

Public Delegate Function ResolveReferenceDelegate(ByVal context As DeserializationContext, ByVal reference As Serializer.ReferenceToResolve) As Boolean

Public Class Serializer

    Private Shared _SpecialChars As Char() = {":"c, ","c, "{"c, "}"c, "["c, "]"c}
    Private Property DefaultSerializers As New List(Of ICustomSerializer)

    Public Property CustomSerializers As New List(Of ICustomSerializer)
    Public Property PropertyBindingFlags As Reflection.BindingFlags = Reflection.BindingFlags.Public Or Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance
    Public Property FieldBindingFlags As Reflection.BindingFlags = Reflection.BindingFlags.GetField Or Reflection.BindingFlags.IgnoreCase Or Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance Or Reflection.BindingFlags.FlattenHierarchy
    Public Property Indent As Integer = 4
    Public Property PrettyFormating As Boolean = True
    Public Property AdvancedSerialization As Boolean = True
    Public Property IgnoreAttributeTypes As Type() = {GetType(Ignore)}
    Public Property SetterAttributeTypes As Type() = {GetType(Setter)}

    Private Const TrueString = "true"
    Private Const FalseString = "false"
    Private Const NullString = "null"

    Public Enum SetterHandlingEnum
        Setter
        FieldAlways
        FieldWhenReadonly
    End Enum

    Public Property SetterHandling As SetterHandlingEnum = SetterHandlingEnum.FieldWhenReadonly

    Private _pseudonyms As Dictionary(Of Type, Dictionary(Of String, String))
    Private _typesProperties As New Dictionary(Of Type, List(Of ComponentModel.PropertyDescriptor))

    Public Sub New()
        Me.DefaultSerializers.Add(New GUIDSerializer)
        Me.DefaultSerializers.Add(New ListSerliazer)
        Me.DefaultSerializers.Add(New DictionarySerializer)
    End Sub

#Region "Events"

    Public Event BeforeSerialize(e As BeforeSerializeEventArgs)

    Friend Sub DoBeforeSerialize(e As BeforeSerializeEventArgs)
        RaiseEvent BeforeSerialize(e)
    End Sub

    Public Event AfterSerializeObject(e As AfterSerializeObjectEventArgs)

    Public Sub DoAfterSerializeObject(e As AfterSerializeObjectEventArgs)
        RaiseEvent AfterSerializeObject(e)
    End Sub

    Public Event HydrateObject(e As HydrateObjectEventArgs)

    Friend Sub DoHydrateObject(e As HydrateObjectEventArgs)
        RaiseEvent HydrateObject(e)
    End Sub

    Public Event TransformObject(e As TransformObjectEventArgs)

    Public Sub DoTransformObject(e As TransformObjectEventArgs)
        RaiseEvent TransformObject(e)
    End Sub

    Public Event OnDeserializedObject(obj As Object)

    Friend Sub DoOnDeserializedObject(obj As Object)
        RaiseEvent OnDeserializedObject(obj)
    End Sub

    Public Event NewInstance(e As NewInstanceEventArgs)

    Friend Sub DoNewInstance(e As NewInstanceEventArgs)
        RaiseEvent NewInstance(e)
    End Sub

    Public Event ObjectDeserialized(e As ObjectDeserializedEventArgs)

    Friend Sub DoObjectDeserialized(e As ObjectDeserializedEventArgs)
        RaiseEvent ObjectDeserialized(e)
    End Sub

#End Region

#Region "Helpers"

    Private Shared Function TryToFindTypeInCollection(context As SerializationContext,
                                                      obj As Object,
                                                      collection As IList(Of ICustomSerializer),
                                                      objectsSerializedOnBranch As IDictionary(Of Object, String),
                                                      ByRef serializer As ICustomSerializer) As Boolean

        For Each ser In collection
            If ser.CanSerialize(context, obj, objectsSerializedOnBranch) Then
                serializer = ser
                Return True
            End If
        Next
        Return False
    End Function

    Private Shared Function GetPseudonymsFor(context As DeserializationContext, t As Type) As IDictionary(Of String, String)

        If context.Serializer._pseudonyms.ContainsKey(t) Then Return context.Serializer._pseudonyms(t)

        Dim resultDict As New Dictionary(Of String, String)

        Dim props = GetPropertiesFor(context, t)
        For Each prop In props
            If resultDict.ContainsKey(prop.Name) Then
                resultDict(prop.Name) = prop.Name
            Else
                resultDict.Add(prop.Name, prop.Name)
            End If

            Dim pseudonymAttrs = (From atr In prop.Attributes Where TypeOf atr Is Pseudonym)
            If pseudonymAttrs.Count > 0 Then
                For Each attr As Pseudonym In pseudonymAttrs
                    If Not resultDict.ContainsKey(attr.Pseudonym) Then
                        resultDict.Add(attr.Pseudonym, prop.Name)
                    End If
                Next
            End If
        Next

        context.Serializer._pseudonyms.Add(t, resultDict)

        Return resultDict

    End Function

    Public Shared Function FromJSonString(str As String) As String

        Dim strlim = str.Length - 1
        Dim sb As New System.Text.StringBuilder()

        For i As Integer = 0 To strlim
            If i > strlim Then Exit For

            Dim c = str(i)
            If Char.Equals(c, "\"c) Then
                Select Case str(i + 1)
                    Case """"c
                        sb.Append("""")
                    Case "\"c
                        sb.Append("\")
                    Case "/"c
                        sb.Append("/")
                    Case "b"c
                        sb.Append(Constants.vbBack)
                    Case "f"c
                        sb.Append(Constants.vbFormFeed)
                    Case "n"c
                        sb.Append(Constants.vbNewLine)
                    Case "r"c
                        sb.Append(Constants.vbCr)
                    Case "t"c
                        sb.Append(Constants.vbTab)
                End Select
                i += 1
            Else
                sb.Append(c)
            End If

        Next

        Return sb.ToString()

    End Function

    Private Shared Function GetPropertiesFor(context As ISerializationContext, type As Type) As List(Of ComponentModel.PropertyDescriptor)
        Dim props As List(Of ComponentModel.PropertyDescriptor) = Nothing
        If Not context.Serializer._typesProperties.TryGetValue(type, props) Then
            'Hyper.ComponentModel.HyperTypeDescriptionProvider.Add(type)
            Dim propsUnsorted = System.ComponentModel.TypeDescriptor.GetProperties(type)
            props = New List(Of ComponentModel.PropertyDescriptor)

            For Each prop As ComponentModel.PropertyDescriptor In propsUnsorted
                Dim ignoreAttr = (From atr In prop.Attributes
                                  From tp In context.Serializer.IgnoreAttributeTypes
                                  Where tp.IsAssignableFrom(atr.GetType())).FirstOrDefault
                ' prop.GetIndexParameters().Count = 0 AndAlso
                If ignoreAttr Is Nothing Then props.Add(prop)
            Next
            context.Serializer._typesProperties.Add(type, props)
        End If
        Return props
    End Function

#End Region

#Region "Serialize"

    Public Sub Serialize(stream As IO.Stream, obj As Object)

        Try
            Dim sw = New FileWriter(stream)
            Dim context As New SerializationContext(Me, sw)
            Serialize(stream, obj, context)
            sw.Flush()
        Catch
            Throw
        End Try

    End Sub

    Public Sub Serialize(stream As IO.Stream, obj As Object, context As SerializationContext)

        Try
            SerializeObject(context, obj, New Dictionary(Of Object, String))
        Catch
            Throw
        End Try

    End Sub

    Private Shared Sub WriteEmptyLine(context As SerializationContext)
        If context.Serializer.PrettyFormating Then context.Stream.WriteLine()
    End Sub

    Private Shared Sub WriteLine(ByVal context As SerializationContext,
                                 ByVal value As String)

        If context.Serializer.PrettyFormating Then
            context.Stream.WriteLine()
            context.Stream.Write(New String(" "c, context.Indent))
        End If

        context.Stream.Write(value)
    End Sub

    Private Shared Sub WriteReference(ByVal context As SerializationContext,
                                      ByVal serializedGuid As Object)

        WriteLine(context, "{")
        context.Indent += context.Serializer.Indent
        WriteLine(context, """$ref"" : """)
        context.Stream.Write(serializedGuid.ToString())
        context.Stream.Write("""")
        context.Indent -= context.Serializer.Indent
        WriteLine(context, "}")
    End Sub

    Private Shared Sub WriteArray(ByVal context As SerializationContext,
                                  ByVal arr As Array,
                                  ByVal objectsSerializedOnBranch As IDictionary(Of Object, String))

        WriteLine(context, "[")
        context.Indent += context.Serializer.Indent

        If arr.Length > 0 Then
            SerializeObject(context, arr.GetValue(0), objectsSerializedOnBranch)

            For i As Integer = 1 To arr.Length - 1
                context.Stream.Write(",")
                SerializeObject(context, arr.GetValue(i), objectsSerializedOnBranch)
            Next
        End If

        context.Indent -= context.Serializer.Indent
        WriteLine(context, "]")
    End Sub

    Private Shared Sub WriteHashTable(ByVal context As SerializationContext,
                                      ByVal list As Hashtable,
                                      ByVal type As Type,
                                      ByVal objectsSerializedOnBranch As IDictionary(Of Object, String))

        WriteLine(context, "{")

        context.Indent += context.Serializer.Indent

        If list.Count > 0 Then

            WriteEmptyLine(context)
            context.Stream.Write(New String(" "c, context.Indent))
            context.Stream.Write("""")
            context.Stream.Write(list.Keys(0).ToString())
            context.Stream.Write(""" : ")
            SerializeObject(context, list.Values(0), objectsSerializedOnBranch)

            For i As Integer = 1 To list.Count - 1
                context.Stream.Write(",")
                WriteEmptyLine(context)
                context.Stream.Write(New String(" "c, context.Indent))
                context.Stream.Write("""")
                context.Stream.Write(list.Keys(i).ToString())
                context.Stream.Write(""" : ")
                SerializeObject(context, list.Values(i), objectsSerializedOnBranch)
            Next
        End If

        context.Indent -= context.Serializer.Indent

        WriteLine(context, "}")
    End Sub

    Private Shared Sub WriteObject(ByVal context As SerializationContext,
                                   ByVal obj As Object,
                                   ByVal type As Type,
                                   ByVal objectsSerializedOnBranch As IDictionary(Of Object, String))

        Dim hashTable As New Hashtable

        If Not objectsSerializedOnBranch.ContainsKey(obj) Then
            Dim objGuid As String = Guid.NewGuid.ToString()

            If context.Serializer.AdvancedSerialization Then
                hashTable.Add("$id", objGuid.ToString())
                hashTable.Add("$type", GetTypeNameFor(type))
            End If

            Dim props = GetPropertiesFor(context, type)
            For i As Integer = 0 To props.Count - 1
                Try
                    Dim e = New BeforeSerializeEventArgs(props(i), type,
                                                        obj, props(i).GetValue(obj),
                                                        context)
                    context.Serializer.DoBeforeSerialize(e)
                    If Not e.Cancel Then hashTable.Add(props(i).Name, e.Val)
                Catch ex As Exception
                    Throw New SerializationException(String.Format("Invalid BeforeSererialize for property {0}", props(i).Name),
                                                     props(i).Name,
                                                     ex)
                End Try
            Next            

            Dim eAfterSerialize = New AfterSerializeObjectEventArgs(obj, type, hashTable, context)
            context.Serializer.DoAfterSerializeObject(eAfterSerialize)

            Dim childObjectsSerializedOnBranch As New Dictionary(Of Object, String)
            For Each kvp In objectsSerializedOnBranch
                childObjectsSerializedOnBranch.Add(kvp.Key, kvp.Value)
            Next
            childObjectsSerializedOnBranch.Add(obj, objGuid)
            WriteHashTable(context, hashTable, type, childObjectsSerializedOnBranch)

            Dim serializedObjDict As Dictionary(Of Object, Object) = Nothing
            If Not context.SerializedObjects.TryGetValue(obj.GetType(), serializedObjDict) Then
                serializedObjDict = New Dictionary(Of Object, Object)()
                context.SerializedObjects.Add(obj.GetType(), serializedObjDict)
            End If
            If Not serializedObjDict.ContainsKey(obj) Then
                serializedObjDict.Add(obj, objGuid)
            End If
        Else
            WriteReference(context, objectsSerializedOnBranch(obj))
        End If
    End Sub

    Public Shared Sub SerializeObject(ByVal context As SerializationContext,
                                      ByVal obj As Object,
                                      ByVal objectsSerializedOnBranch As IDictionary(Of Object, String))
        If obj Is Nothing Then
            context.Stream.Write(NullString)
            Return
        End If

        Dim type = obj.GetType()

        Dim serializer As ICustomSerializer = Nothing

        If TypeOf obj Is String OrElse TypeOf obj Is Char Then
            context.Stream.Write("""")
            context.Stream.WriteString(obj.ToString())
            context.Stream.Write("""")

        ElseIf TypeOf obj Is Boolean Then
            context.Stream.Write(If(CBool(obj), TrueString, FalseString))

        ElseIf TypeOf obj Is DateTime Then
            context.Stream.Write("""")
            context.Stream.Write(DirectCast(obj, DateTime).ToString("yyyy/MM/dd HH:mm:ss.ffff"))
            context.Stream.Write("""")

        ElseIf Array.IndexOf({GetType(Int32), GetType(Int64), GetType(Double), GetType(Int16), GetType(Single), GetType(Decimal), GetType(Byte), GetType(SByte), GetType(UInt32), GetType(UInt64), GetType(UInt16)}, type) > -1 Then
            context.Stream.Write(DirectCast(obj, IConvertible).ToString(Globalization.CultureInfo.InvariantCulture.NumberFormat))

        ElseIf type.IsEnum Then
            context.Stream.Write(CInt(obj).ToString(Globalization.CultureInfo.InvariantCulture.NumberFormat))

        ElseIf TypeOf obj Is Byte() Then
            context.Stream.Write("""")
            context.Stream.WriteString(Convert.ToBase64String(DirectCast(obj, Byte())))
            context.Stream.Write("""")

        ElseIf type.IsArray() Then
            WriteArray(context, DirectCast(obj, Array), objectsSerializedOnBranch)

        ElseIf TypeOf obj Is Hashtable Then
            WriteHashTable(context, DirectCast(obj, Hashtable), type, objectsSerializedOnBranch)

        Else

            Dim serializedGuid As Object = Nothing

            If context.Serializer.AdvancedSerialization Then
                Dim serializedObjDict As Dictionary(Of Object, Object) = Nothing
                If context.SerializedObjects.TryGetValue(obj.GetType, serializedObjDict) Then
                    serializedGuid = (From so In serializedObjDict
                                      Where Object.Equals(so.Key, obj)
                                      Select so.Value).FirstOrDefault
                End If
            End If

            If context.Serializer.AdvancedSerialization AndAlso serializedGuid IsNot Nothing Then
                WriteReference(context, serializedGuid)

            ElseIf TryToFindTypeInCollection(context, obj, context.Serializer.CustomSerializers, objectsSerializedOnBranch, serializer) Then
                serializer.SerializeObject(context, obj, objectsSerializedOnBranch)

            ElseIf TryToFindTypeInCollection(context, obj, context.Serializer.DefaultSerializers, objectsSerializedOnBranch, serializer) Then
                serializer.SerializeObject(context, obj, objectsSerializedOnBranch)

            Else
                WriteObject(context, obj, type, objectsSerializedOnBranch)

            End If
        End If
    End Sub

#End Region

#Region "Deserialize"

    Public Function Deserialize(Of T)(stream As IO.Stream) As T

        Dim resultObj As Object

        Dim sr As IO.StreamReader = Nothing
        Try
            sr = New IO.StreamReader(stream)

            ' Initialization
            Dim context = New DeserializationContext(Me, sr)
            _pseudonyms = New Dictionary(Of Type, Dictionary(Of String, String))

            ' And actual Deserialization
            resultObj = DeserializeStream(context, GetType(T), Nothing, Nothing)
            If TypeOf resultObj Is Hashtable Then
                resultObj = TransformHashTable(DirectCast(resultObj, Hashtable), GetType(T), GetType(T) IsNot GetType(Object), Nothing, Nothing, context)
            End If

            ' Resolve previously unresolved references (if possible)
            Dim lastUnresolvedReferencesCount As Integer

            While (From r In context.ReferencesToResolve Where Not r.HasBeenResolved).Count > 0 And lastUnresolvedReferencesCount <> context.ReferencesToResolve.Count
                lastUnresolvedReferencesCount = (From r In context.ReferencesToResolve Where Not r.HasBeenResolved).Count

                For Each r2r In context.ReferencesToResolve
                    If Not r2r.HasBeenResolved Then
                        r2r.HasBeenResolved = r2r.ResolveReference(context, r2r)
                    End If
                Next
            End While
        Catch
            Throw
        Finally
            If sr IsNot Nothing Then sr.Close()
        End Try

        Return DirectCast(resultObj, T)

    End Function

    Public Function Deserialize(ByVal stream As IO.Stream) As Object

        Return Me.Deserialize(Of Object)(stream)

    End Function

#Region "Convenience converters"

    Public Function DeserializeInt(stream As IO.Stream) As Integer

        Return Integer.Parse(Deserialize(stream).ToString())

    End Function

    Public Function DeserializeDateTime(stream As IO.Stream) As DateTime

        Return CDate(ParseDateTime(Deserialize(stream)))

    End Function

    Public Function DeserializeDouble(stream As IO.Stream) As Double

        Return Double.Parse(Deserialize(stream).ToString())

    End Function

    Public Function DeserializeDecimal(stream As IO.Stream) As Decimal

        Return Decimal.Parse(Deserialize(stream).ToString())

    End Function

    Public Function DeserializeSingle(stream As IO.Stream) As Single

        Return Single.Parse(Deserialize(stream).ToString())

    End Function

    Public Function DeserializeBoolean(stream As IO.Stream) As Boolean

        Return Boolean.Parse(Deserialize(stream).ToString())

    End Function

    Public Function DeserializeString(stream As IO.Stream) As String

        Return Deserialize(stream).ToString()

    End Function

#End Region

    Private Shared Sub ReadUntilNextDelimiter(ByVal sr As IO.StreamReader, ByRef value As String, ByRef delimiter As String)

        Dim sb As New System.Text.StringBuilder()
        Dim prevChr As Char
        Dim chr As Char

        Dim openQuote = False
        Dim escapeChar = False

        Do
            chr = ChrW(sr.Read())
            If openQuote OrElse Array.IndexOf(_SpecialChars, chr) = -1 Then
                sb.Append(chr)
            End If
            escapeChar = (prevChr = "\")
            prevChr = chr
            If chr = """" AndAlso Not escapeChar Then openQuote = Not openQuote
        Loop While Not sr.EndOfStream AndAlso (openQuote OrElse (Not escapeChar AndAlso Array.IndexOf(_SpecialChars, chr) = -1))

        value = sb.ToString().Trim()
        delimiter = chr

    End Sub

    Public Shared Function TransformHashTable(ByVal hashTable As Hashtable,
                                              ByVal targetType As Type, forceTargetType As Boolean,
                                              ByVal obj As Object, prop As System.ComponentModel.PropertyDescriptor,
                                              ByVal context As DeserializationContext) As Object

        Dim resultObj As Object = hashTable

        ' Return Nothing is the magic key that this object reference has not yet been found
        ' and deserialized. Calling code should increment context.ReferencesToResolve.
        If hashTable.ContainsKey("$ref") Then
            If context.SerializedObjects.ContainsKey(hashTable("$ref").ToString()) Then
                Return context.SerializedObjects(hashTable("$ref").ToString()).Result
            Else
                Return Nothing
            End If
        End If

        ' foreceTargetType = True means we're being called from SerializeObject (or custom code)
        ' and want to get the type from the supplied type [Deserialize(Of T)] instead of the JSON text.
        Dim resultType As Type, typeName As String
        If Not forceTargetType AndAlso hashTable.ContainsKey("$type") Then
            typeName = hashTable("$type").ToString()
            resultType = Type.GetType(typeName)
        ElseIf targetType Is GetType(Object) Then
            Return hashTable
        Else
            typeName = targetType.AssemblyQualifiedName
            resultType = targetType
        End If

        ' If this object is recognised by any of the known serializers, return that
        resultObj = TryCustomDeserialize(context, typeName, resultType, hashTable, obj, prop)
        If resultObj IsNot Nothing Then Return resultObj

        ' Calling code should know how to handle the fact that the object has come back unchanged.
        If resultType Is Nothing Then Return hashTable

        ' Finally we can rest assured this is an object we want to deserialize ourselves
        Dim e As New NewInstanceEventArgs(context, resultType, hashTable)
        context.Serializer.DoNewInstance(e)
        Try
            If e.Handled Then
                resultObj = e.NewObj
            Else
                resultObj = System.Runtime.Serialization.FormatterServices.GetUninitializedObject(resultType)
            End If
        Catch ex As Exception
            Throw New SerializationException(String.Format("Error creating type {0}", resultType.Name),
                                             ex)
        End Try

        Dim pseudonyms = GetPseudonymsFor(context, resultType)

        Dim allPropertiesForType = GetPropertiesFor(context, resultType)

        For i As Integer = 0 To hashTable.Count - 1
            ' Retrieve the relevant property on the targetType
            Dim propNameDes = hashTable.Keys(i).ToString()
            Dim propName As String = String.Empty, propExst As System.ComponentModel.PropertyDescriptor = Nothing
            If pseudonyms.TryGetValue(propNameDes, propName) Then
                propExst = (From p2 In allPropertiesForType Where p2.Name = propName).FirstOrDefault
            End If

            If propExst IsNot Nothing Then
                Try
                    SetProperty(context, resultObj, propExst, hashTable.Values(i))
                Catch ex As Exception
                    Throw New SerializationException(String.Format("Error deserializing property {0}", propExst.Name),
                                                     propExst.Name,
                                                     ex)
                End Try
            End If
        Next

        If hashTable.ContainsKey("$id") Then
            context.SerializedObjects.Add(hashTable("$id").ToString(), New DeserializedObject(resultObj, hashTable))
            context.Serializer.DoOnDeserializedObject(resultObj)
        End If

        Return resultObj

    End Function

    Friend Shared Function TryTransformHashTable(ByVal context As DeserializationContext,
                                                 ByVal resultObj As Object, ByVal prop As System.ComponentModel.PropertyDescriptor,
                                                 ByVal val As Object, ByVal targetType As Type) As Object

        Dim eVal = TransformHashTable(DirectCast(val, Collections.Hashtable), targetType, False, resultObj, prop, context)
        ' eVal is Nothing means that this is a reference which has not yet been resolved. Store for later.
        If eVal Is Nothing AndAlso DirectCast(val, Collections.Hashtable).ContainsKey("$ref") Then
            ' Add a reference for later
            Try
                context.ReferencesToResolve.Add(New ReferenceToResolve(resultObj,
                                                                       prop,
                                                                       DirectCast(val, Collections.Hashtable)("$ref").ToString(),
                                                                       -1,
                                                                       targetType,
                                                                       AddressOf DefaultReferenceResolver))
            Catch ex As Exception
                Throw New SerializationException(String.Format("Could not store reference for {0}", DirectCast(val, Collections.Hashtable)("$ref").ToString()),
                                                 ex)
            End Try

        End If
        Return eVal

    End Function

    Private Shared Function TryCustomDeserialize(ByVal context As DeserializationContext,
                                                 ByVal typeName As String, ByVal resultType As Type,
                                                 ByVal val As Object,
                                                 ByVal obj As Object,
                                                 ByVal prop As ComponentModel.PropertyDescriptor) As Object

        For Each customSerializer In context.Serializer.CustomSerializers
            If customSerializer.CanDeserialize(context, typeName, val) Then
                Dim result = customSerializer.DeserializeObject(context, resultType, val, obj, prop)
                If result IsNot Nothing Then Return result
            End If
        Next

        For Each defaultSerializer In context.Serializer.DefaultSerializers
            If defaultSerializer.CanDeserialize(context, typeName, val) Then
                Dim result = defaultSerializer.DeserializeObject(context, resultType, val, obj, prop)
                If result IsNot Nothing Then Return result
            End If
        Next

        Return Nothing
    End Function

    Private Shared Sub SetProperty(ByVal context As DeserializationContext,
                                   ByRef resultObj As Object,
                                   ByVal prop As System.ComponentModel.PropertyDescriptor,
                                   ByVal val As Object)

        Dim propType = prop.PropertyType

        If propType.IsGenericType AndAlso propType.GetGenericTypeDefinition() Is GetType(Nullable(Of )) Then
            propType = prop.PropertyType.GetGenericArguments().FirstOrDefault
        End If

        If val Is Nothing Then
            ' Save the object as is

        ElseIf propType Is GetType(Integer) Then
            val = CInt(ParseNumber(val))

        ElseIf propType Is GetType(Single) Then
            val = CSng(ParseNumber(val))

        ElseIf propType Is GetType(Int16) Then
            val = CShort(ParseNumber(val))

        ElseIf propType Is GetType(Byte) Then
            val = CByte(ParseNumber(val))

        ElseIf propType Is GetType([Enum]) Then
            val = CInt(ParseNumber(val))

        ElseIf propType Is GetType(Double) Then
            val = CDbl(ParseNumber(val))

        ElseIf propType Is GetType(Long) Then
            val = CLng(ParseNumber(val))

        ElseIf propType Is GetType(Decimal) Then
            val = CDec(ParseNumber(val))

        ElseIf propType Is GetType(Boolean) Then
            val = CBool(ParseBoolean(val))

        ElseIf propType.IsEnum Then
            val = CInt(ParseNumber(val))

        ElseIf propType Is GetType(Byte()) Then
            val = Convert.FromBase64String(CStr(val))

        ElseIf propType Is GetType(Char) Then
            val = CChar(val)

        ElseIf propType Is GetType(DateTime) Then
            val = CDate(ParseDateTime(val))

        ElseIf propType Is GetType(String) Then
            val = val.ToString()

        ElseIf TypeOf val Is Hashtable Then
            Dim eVal = TryTransformHashTable(context, resultObj, prop, val, propType)
            ' eVal is Nothing means that this is a reference which has not yet been resolved. Store for later.
            If eVal IsNot Nothing AndAlso propType.IsAssignableFrom(eVal.GetType()) Then
                val = eVal
            End If

        Else
            ' If this object is recognised by any of the known serializers, return that
            Dim result = TryCustomDeserialize(context, propType.AssemblyQualifiedName, propType, val, resultObj, prop)
            If result IsNot Nothing Then val = result

        End If

        Dim eh As New HydrateObjectEventArgs(resultObj, val, prop, context)
        context.Serializer.DoHydrateObject(eh)
        If Not eh.Handled AndAlso (val Is Nothing OrElse propType.IsValueType OrElse propType.IsAssignableFrom(val.GetType())) Then
            Dim e As New TransformObjectEventArgs(resultObj, val)
            context.Serializer.DoTransformObject(e)
            If e.Handled Then val = e.Val

            ' The settter mode may be set on the property
            Dim setterHandling = context.Serializer.SetterHandling

            Dim setterAttr = (From atr In prop.Attributes
                              From tp In context.Serializer.SetterAttributeTypes
                              Where tp.IsAssignableFrom(atr.GetType())
                              Select atr).FirstOrDefault

            If setterAttr IsNot Nothing Then
                Dim setProp = setterAttr.GetType().GetProperty("SetterHandling")
                If setProp IsNot Nothing Then
                    setterHandling = DirectCast(setProp.GetValue(setterAttr, Nothing), SetterHandlingEnum)
                End If
            End If

            ' Allow for the field to be set directly instead of the Setter when possible
            Select Case setterHandling 
                Case SetterHandlingEnum.Setter
                    If Not prop.IsReadOnly Then
                        SetPropertyValue(prop, resultObj, val)
                    End If

                Case SetterHandlingEnum.FieldAlways
                    SetFieldValue(prop, resultObj, val, context.Serializer.FieldBindingFlags)

                Case SetterHandlingEnum.FieldWhenReadonly
                    If Not prop.IsReadOnly Then
                        SetPropertyValue(prop, resultObj, val)
                    Else
                        SetFieldValue(prop, resultObj, val, context.Serializer.FieldBindingFlags)
                    End If
            End Select
        End If
    End Sub

    Private Shared Sub SetPropertyValue(prop As ComponentModel.PropertyDescriptor, obj As Object, val As Object)
        prop.SetValue(obj, val)
    End Sub

    Private Shared Function GetField(t As Type, fieldName As String, fieldBindingFlags As Reflection.BindingFlags) As Reflection.FieldInfo
        If t Is Nothing Then Return Nothing
        Dim field = t.GetField(fieldName, fieldBindingFlags)
        If field Is Nothing Then
            field = GetField(t.BaseType, fieldName, fieldBindingFlags)
        End If
        Return field
    End Function

    Private Shared Sub SetFieldValue(prop As ComponentModel.PropertyDescriptor, obj As Object, val As Object, fieldBindingFlags As Reflection.BindingFlags)
        Dim fld = GetField(obj.GetType(), "_" + prop.Name, fieldBindingFlags)
        If fld IsNot Nothing Then
            fld.SetValue(obj, val)
        End If
    End Sub

    Private Shared Function DefaultReferenceResolver(ByVal context As DeserializationContext, ByVal reference As ReferenceToResolve) As Boolean
        If context.SerializedObjects.ContainsKey(reference.Reference) Then
            Dim transformedReference As Object
            If reference.PropInfo.PropertyType.IsAssignableFrom(context.SerializedObjects(reference.Reference).Result.GetType()) Then
                transformedReference = context.SerializedObjects(reference.Reference).Result
            Else
                transformedReference = TransformHashTable(context.SerializedObjects(reference.Reference).Origin, reference.PropInfo.PropertyType, True, reference.Obj, reference.PropInfo, context)
            End If

            SetProperty(context,
                        reference.Obj,
                        reference.PropInfo,
                        transformedReference)

            Return True
        Else
            Return False
        End If
    End Function

    Private Shared Function DeserializeStream(ByVal context As DeserializationContext,
                                              ByVal targetType As Type,
                                              ByVal value As String, ByRef delimiter As String) As Object

        If value Is Nothing Then
            ReadUntilNextDelimiter(context.Stream, value, delimiter)
        End If

        Select Case delimiter
            Case "{"c
                Return DeserializeObject(context, targetType, value, delimiter)

            Case "["c
                Return DeserializeArray(context, targetType, value, delimiter)

            Case ":"c
                Return DeserializeValuePair(context, targetType, value, delimiter)

            Case "", ","c
                Return DeserializeValue(context, targetType, value, delimiter)

            Case Else
                Return DeserializeValue(context, targetType, value, delimiter)
        End Select

    End Function

    Private Shared Function DeserializeArray(ByVal context As DeserializationContext,
                                             ByVal targetType As Type,
                                             ByVal value As String, ByRef delimiter As String) As Array

        ' First read all the data from the JSON text.
        Dim arrayList As New ArrayList()

        Do

            ReadUntilNextDelimiter(context.Stream, value, delimiter)

            Dim val = DeserializeStream(context, targetType, value, delimiter)

            If Not (delimiter = "]" AndAlso TypeOf val Is String AndAlso String.IsNullOrEmpty(CStr(val))) Then
                arrayList.Add(val)
            End If

        Loop Until delimiter = "]"

        ReadUntilNextDelimiter(context.Stream, value, delimiter)

        ' Next try to interpret the data.
        If targetType Is Nothing Then
            Return arrayList.ToArray()
        Else
            If targetType Is GetType(Object) OrElse Not targetType.IsArray() Then
                targetType = GetType(Object())
            End If

            Dim arrInst = DirectCast(Activator.CreateInstance(targetType, arrayList.Count), Array)
            For j As Integer = 0 To arrayList.Count - 1
                Try
                    Dim val = arrayList(j)
                    If Not targetType.GetElementType() Is GetType(Object) AndAlso TypeOf val Is Hashtable Then
                        val = TryTransformHashTable(context, arrInst, Nothing, val, targetType.GetElementType())

                        ' val is Nothing means that this is a reference which has not yet been resolved. Store for later.
                        If val IsNot Nothing Then arrInst.SetValue(val, j)
                    Else
                        arrInst.SetValue(val, j)
                    End If
                Catch 
                    Throw
                End Try
            Next

            Return arrInst
        End If

    End Function

    Private Shared Function DeserializeObject(ByVal context As DeserializationContext,
                                              ByVal targetType As Type,
                                              ByVal value As String, ByRef delimiter As String) As Hashtable

        Dim hashTable As New Hashtable()

        ReadUntilNextDelimiter(context.Stream, value, delimiter)

        Do

            If Not delimiter = "}"c Then
                Dim key = DeserializeValue(context, GetType(Object), value, delimiter).ToString().Trim(""""c)

                ReadUntilNextDelimiter(context.Stream, value, delimiter)

                Dim valProp = targetType.GetProperty(key)
                Dim valType As Type
                If valProp IsNot Nothing Then
                    valType = valProp.PropertyType
                Else
                    valType = GetType(Object)
                End If
                Dim val = DeserializeStream(context, valType, value, delimiter)

                Try
                    hashTable.Add(key, val)
                Catch ex As Exception
                    Throw New SerializationException("Could not add key to hashtable: " + If(key IsNot Nothing, key.ToString(), String.Empty), ex)
                End Try

                If delimiter = "," Then
                    ReadUntilNextDelimiter(context.Stream, value, delimiter)
                End If
            End If

        Loop Until delimiter = "}"

        If Not hashTable.ContainsKey("$ref") Then
            context.Serializer.DoObjectDeserialized(New ObjectDeserializedEventArgs(context, hashTable))
        End If

        ReadUntilNextDelimiter(context.Stream, value, delimiter)

        Return hashTable

    End Function

    Private Shared Function ParseNumber(ByVal value As Object) As Object

        Dim dbl As Double

        If value IsNot Nothing AndAlso
            Double.TryParse(value.ToString(),
                            Globalization.NumberStyles.Float,
                            Globalization.CultureInfo.InvariantCulture,
                            dbl) Then
            Return dbl
        Else
            Return value
        End If

    End Function

    Private Shared Function ParseBoolean(ByVal value As Object) As Object

        Dim bool As Boolean

        If value IsNot Nothing AndAlso
            Boolean.TryParse(value.ToString(), bool) Then
            Return bool
        Else
            Return value
        End If

    End Function

    Private Shared Function ParseDateTime(ByVal value As Object) As Object

        Dim dte As DateTime

        If value IsNot Nothing AndAlso
            DateTime.TryParse(value.ToString().Trim(""""c),
                              Globalization.CultureInfo.InvariantCulture,
                              Globalization.DateTimeStyles.AssumeLocal,
                              dte) Then
            Return dte
        Else
            Return value
        End If

    End Function

    Private Shared Function DeserializeValue(ByVal context As DeserializationContext,
                                             ByVal targetType As Type,
                                             ByVal value As String, ByRef delimiter As String) As Object
        If value.ToLower = NullString Then
            Return Nothing
        Else
            Return FromJSonString(value.ToString()).Trim(""""c)
        End If

    End Function

    Private Shared Function DeserializeValuePair(ByVal context As DeserializationContext,
                                                 ByVal targetType As Type,
                                                 ByVal value As String, ByRef delimiter As String) As KeyValuePair(Of Object, Object)

        Dim genericTypeArguments = targetType.GetGenericArguments
        Dim genericTypeArgument As Type
        If genericTypeArguments.Count > 0 Then
            genericTypeArgument = genericTypeArguments(0)
        Else
            genericTypeArgument = GetType(Object)
        End If
        Dim key = DeserializeStream(context, genericTypeArgument, value, "")
        ReadUntilNextDelimiter(context.Stream, value, delimiter)
        Dim val = DeserializeStream(context, genericTypeArgument, value, delimiter)

        Return New KeyValuePair(Of Object, Object)(key, val)

    End Function

#End Region

#Region "ReferenceToResolve"

    Public Class ReferenceToResolve
        Public Property Obj As Object
        Public Property PropInfo As System.ComponentModel.PropertyDescriptor
        Public Property Reference As Object
        Public Property Index As Integer = -1
        Public Property TargetType As Type
        Public Property ResolveReference As ResolveReferenceDelegate
        Public Property HasBeenResolved As Boolean = False

        Public Sub New(ByVal obj As Object,
                       ByVal propInfo As System.ComponentModel.PropertyDescriptor,
                       ByVal reference As Object,
                       ByVal index As Integer,
                       ByVal targetType As Type,
                       ByVal resolveReference As ResolveReferenceDelegate)
            Me.Obj = obj
            Me.PropInfo = propInfo
            Me.Reference = reference
            Me.Index = index
            Me.TargetType = targetType
            Me.ResolveReference = resolveReference
        End Sub
    End Class

#End Region

End Class