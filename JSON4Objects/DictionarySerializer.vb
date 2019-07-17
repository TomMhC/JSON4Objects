Friend Class DictionarySerializer
    Implements ICustomSerializer

    Public Function ResolveReference(ByVal context As DeserializationContext, ByVal reference As Serializer.ReferenceToResolve) As Boolean

        Dim transformedKey = Nothing, transformedVal = Nothing

        Dim obj = DirectCast(reference.Reference, Hashtable)
        If TypeOf obj.Keys(reference.Index) Is Hashtable AndAlso DirectCast(obj.Keys(reference.Index), Hashtable).ContainsKey("$ref") Then
            If context.SerializedObjects.ContainsKey(DirectCast(obj.Keys(reference.Index), Hashtable)("$ref").ToString()) Then
                transformedKey = context.SerializedObjects(DirectCast(obj.Keys(reference.Index), Hashtable)("$ref").ToString()).Result
            Else
                Return False
            End If
        Else
            transformedKey = Serializer.FromJSonString(obj.Keys(reference.Index).ToString())
        End If

        If TypeOf obj.Values(reference.Index) Is Hashtable AndAlso DirectCast(obj.Values(reference.Index), Hashtable).ContainsKey("$ref") Then
            If context.SerializedObjects.ContainsKey(DirectCast(obj.Values(reference.Index), Hashtable)("$ref").ToString()) Then
                transformedVal = context.SerializedObjects(DirectCast(obj.Values(reference.Index), Hashtable)("$ref").ToString()).Result
            Else
                Return False
            End If
        Else
            transformedVal = Serializer.FromJSonString(obj.Values(reference.Index).ToString())
        End If

        DirectCast(reference.Obj, IDictionary).Add(transformedKey, transformedVal)

        Return True

    End Function

    Public Function DeserializeObject(ByVal context As DeserializationContext, targetType As Type, val As Object,
                                      obj As Object, prop As System.ComponentModel.PropertyDescriptor) As Object Implements ICustomSerializer.DeserializeObject

        Dim genericKeyType As Type = Nothing, genericValueType As Type = Nothing

        Dim resultObj As IDictionary

        Dim hashTable = DirectCast(val, Hashtable)

        If hashTable.ContainsKey("$type") Then
            Dim resultType = Type.GetType(hashTable("$type").ToString())
            targetType = resultType
        End If

        If targetType.IsGenericType Then
            genericKeyType = targetType.GetGenericArguments(0)
            genericValueType = targetType.GetGenericArguments(1)
        End If

        resultObj = DirectCast(Activator.CreateInstance(targetType, True), IDictionary)

        For i As Integer = 0 To hashTable.Count - 1
            If Not hashTable.Keys(i).ToString() = "$type" Then
                Dim transformedKey = Nothing, transformedVal = Nothing

                If TypeOf hashTable.Keys(i) Is Hashtable Then
                    transformedKey = Serializer.TransformHashTable(DirectCast(hashTable.Keys(i), Collections.Hashtable), targetType.GetGenericArguments(0), False, obj, prop, context)

                    Dim e As New TransformObjectEventArgs(resultObj, transformedKey)
                    context.Serializer.DoTransformObject(e)
                    If e.Handled Then transformedKey = e.Val
                ElseIf genericKeyType IsNot Nothing Then
                    Dim key = hashTable.Keys(i).ToString()
                    transformedKey = Serializer.TransformDeserializedString(context, Nothing, genericKeyType, Nothing, key)
                Else
                    Dim key = hashTable.Keys(i).ToString()
                    transformedKey = Serializer.FromJSonString(key)
                End If

                If TypeOf hashTable.Values(i) Is Hashtable Then
                    transformedVal = Serializer.TransformHashTable(DirectCast(hashTable.Values(i), Collections.Hashtable), targetType.GetGenericArguments(1), False, obj, prop, context)

                    Dim e As New TransformObjectEventArgs(resultObj, transformedVal)
                    context.Serializer.DoTransformObject(e)
                    If e.Handled Then transformedVal = e.Val
                Else
                    Dim hval = hashTable.Values(i)
                    If genericValueType IsNot Nothing Then
                        transformedVal = Serializer.TransformDeserializedString(context, Nothing, genericValueType, Nothing, hval)
                    ElseIf TypeOf hval Is String Then
                        transformedVal = Serializer.FromJSonString(hval.ToString().Trim(""""c))
                    Else
                        transformedVal = hval
                    End If
                End If

                If transformedKey Is Nothing OrElse transformedVal Is Nothing Then
                    context.ReferencesToResolve.Add(New Serializer.ReferenceToResolve(resultObj,
                                                                                      Nothing,
                                                                                      hashTable,
                                                                                      i,
                                                                                      targetType,
                                                                                      AddressOf ResolveReference))
                Else
                    resultObj.Add(transformedKey, transformedVal)
                End If
            End If
        Next

        Return resultObj

    End Function

    Public Sub SerializeObject(ByVal context As SerializationContext, ByVal obj As Object, objectsSerializedOnBranch As IDictionary(Of Object, String)) Implements ICustomSerializer.SerializeObject

        Dim iDict = DirectCast(obj, IDictionary)

        Dim htDict As New Hashtable()

        If context.Serializer.AdvancedSerialization Then
            htDict.Add("$type", GetTypeNameFor(obj.GetType))
        End If

        For i As Integer = 0 To iDict.Count - 1
            htDict.Add(iDict.Keys(i), iDict.Values(i))
        Next

        Serializer.SerializeObject(context, htDict, objectsSerializedOnBranch)

    End Sub

    Public Function CanSerialize(ByVal context As SerializationContext, ByVal obj As Object, objectsSerializedOnBranch As IDictionary(Of Object, String)) As Boolean Implements ICustomSerializer.CanSerialize
        Return TypeOf obj Is IDictionary
    End Function

    Public Function CanDeserialize(ByVal context As DeserializationContext, ByVal typeName As String, obj As Object) As Boolean Implements ICustomSerializer.CanDeserialize
        Return GetType(IDictionary).IsAssignableFrom(Type.GetType(typeName))
    End Function
End Class