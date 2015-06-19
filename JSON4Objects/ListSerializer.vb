Friend Class ListSerliazer
    Implements ICustomSerializer

    Public Function ResolveReference(ByVal context As DeserializationContext, ByVal reference As Serializer.ReferenceToResolve) As Boolean
        If context.SerializedObjects.ContainsKey(reference.Reference) Then
            Dim val = context.SerializedObjects(reference.Reference).Result
            DirectCast(reference.Obj, IList).Insert(reference.Index, val)
            Return True
        Else
            Return False
        End If
    End Function

    Public Function DeserializeObject(ByVal context As DeserializationContext, targetType As Type, val As Object,
                                      obj As Object, prop As ComponentModel.PropertyDescriptor) As Object Implements ICustomSerializer.DeserializeObject

        For Each item In DirectCast(val, IList)
            If TypeOf item Is Hashtable AndAlso DirectCast(item, Hashtable).Count = 1 AndAlso DirectCast(item, Hashtable).ContainsKey("$type") Then
                targetType = Type.GetType(DirectCast(item, Hashtable)("$type").ToString())
                Exit For
            End If
        Next

        Dim resultObj = DirectCast(Activator.CreateInstance(targetType), IList)

        For Each item In DirectCast(val, IList)
            If Not (TypeOf item Is Hashtable AndAlso DirectCast(item, Hashtable).Count = 1 AndAlso DirectCast(item, Hashtable).ContainsKey("$type")) Then
                Dim itemType As Type = Nothing
                If targetType.IsGenericType Then
                    itemType = targetType.GetGenericArguments(0)
                Else
                    itemType = GetType(Object)
                End If

                If TypeOf item Is Hashtable Then
                    Dim transformedObj = Serializer.TransformHashTable(DirectCast(item, Hashtable), itemType, False, obj, prop, context)
                    If transformedObj IsNot Nothing Then
                        Dim e As New TransformObjectEventArgs(resultObj, transformedObj)
                        context.Serializer.DoTransformObject(e)
                        If e.Handled Then transformedObj = e.Val

                        resultObj.Add(transformedObj)
                    Else
                        context.ReferencesToResolve.Add(New Serializer.ReferenceToResolve(resultObj,
                                                                                          Nothing,
                                                                                          DirectCast(item, Collections.Hashtable)("$ref").ToString(),
                                                                                          resultObj.Count,
                                                                                          itemType,
                                                                                          AddressOf ResolveReference))
                    End If
                Else
                    resultObj.Add(item)
                End If
            End If
        Next

        Return resultObj

    End Function

    Public Sub SerializeObject(ByVal context As SerializationContext, ByVal obj As Object, objectsSerializedOnBranch As IDictionary(Of Object, String)) Implements ICustomSerializer.SerializeObject

        Dim iList = DirectCast(obj, IList)

        Dim listLength = If(context.Serializer.AdvancedSerialization, iList.Count, iList.Count - 1)
        Dim arItems(listLength) As Object

        If context.Serializer.AdvancedSerialization Then
            Dim htList As New Hashtable()

            If context.Serializer.AdvancedSerialization Then
                htList.Add("$type", GetTypeNameFor(obj.GetType))
            End If

            arItems(0) = htList
        End If

        Dim advancedSerializationAdd = If(context.Serializer.AdvancedSerialization, 1, 0)
        For i As Integer = 0 To iList.Count - 1
            arItems(i + advancedSerializationAdd) = iList(i)
        Next

        Serializer.SerializeObject(context, arItems, objectsSerializedOnBranch)

    End Sub

    Public Function CanSerialize(ByVal context As SerializationContext, ByVal obj As Object, objectsSerializedOnBranch As IDictionary(Of Object, String)) As Boolean Implements ICustomSerializer.CanSerialize
        Return TypeOf obj Is IList
    End Function

    Public Function CanDeserialize(ByVal context As DeserializationContext, ByVal typeName As String, obj As Object) As Boolean Implements ICustomSerializer.CanDeserialize
        Dim targetType = Type.GetType(typeName)
        If targetType Is Nothing OrElse targetType.IsArray() Then Return False
        Return GetType(IList).IsAssignableFrom(targetType) OrElse (targetType.IsGenericType AndAlso targetType.GetGenericTypeDefinition() = GetType(IList(Of )))
    End Function
End Class