Public Module Helpers

    ''' <summary>
    ''' Returns the full typename, even for generics
    ''' </summary>
    Public Function GetTypeNameFor(t As Type) As String
        Dim args = t.GetGenericArguments()
        If args.Count = 0 Then
            Return _
                t.Namespace +
                "." + If(t.DeclaringType IsNot Nothing, $"{t.DeclaringType.Name}+", String.Empty) + t.Name +
                ", " + t.Assembly.GetName().Name
        Else
            Dim sb As New System.Text.StringBuilder

            If Not String.IsNullOrEmpty(t.Namespace) Then
                ' Namespace may be empty for anonymous types
                sb.Append(t.Namespace + ".")
            End If

            sb.Append(t.Name + "[")

            sb.Append("[" + GetTypeNameFor(args(0)) + "]")

            For i As Integer = 1 To args.Count - 1
                sb.Append(", [" + GetTypeNameFor(args(i)) + "]")
            Next

            sb.Append("], " + t.Assembly.GetName().Name)

            Return sb.ToString()
        End If
    End Function

    Const Back = CChar(Constants.vbBack)
    Const FormFeed = CChar(Constants.vbFormFeed)
    Const NewLine = CChar(Constants.vbNewLine)
    Const Cr = CChar(Constants.vbCr)
    Const Tab = CChar(Constants.vbTab)

    ''' <summary>
    ''' Use this method to write a JSON escaped string to a StringBuilder object
    ''' </summary>
    Public Sub WriteJSonString(sb As System.Text.StringBuilder, str As String)

        For i As Integer = 0 To str.Length - 1

            Dim c = str(i)

            Select Case c

                Case "\"c
                    sb.Append("\\")
                Case """"c
                    sb.Append("\""")
                Case "/"c
                    sb.Append("\/")
                Case Back
                    sb.Append("\b")
                Case FormFeed
                    sb.Append("\f")
                Case NewLine
                    sb.Append("\n")
                Case Cr
                    sb.Append("\r")
                Case Tab
                    sb.Append("\t")
                Case Else
                    sb.Append(c)


            End Select

        Next

    End Sub

    ''' <summary>
    ''' Use this method to create a JSON escaped string
    ''' </summary>
    Public Function ToJSonString(str As String) As String
        Dim sb As New System.Text.StringBuilder(str)

        WriteJSonString(sb, str)

        Return sb.ToString()
    End Function

End Module
