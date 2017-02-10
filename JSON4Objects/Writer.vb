Public Interface IWriter

    Sub WriteString(val As String)
    Sub Write(val As String)
    Sub WriteLine()
    Sub Flush()
    Property Stream As IO.Stream

End Interface

Public Class FileWriter
    Implements IWriter

    Private _sb As New Text.StringBuilder()
    Private _streamWriter As IO.StreamWriter

    Public Property Stream As IO.Stream Implements IWriter.Stream

    Public Sub New(stream As IO.Stream)
        _stream = stream
        If stream IsNot Nothing Then _streamWriter = New IO.StreamWriter(_Stream, System.Text.Encoding.UTF8)
    End Sub

    Public Sub Write(val As String) Implements IWriter.Write
        If val.Length > 10000 Then
            If _streamWriter IsNot Nothing Then _streamWriter.Write(_sb.ToString())
            If _streamWriter IsNot Nothing Then _streamWriter.Write(val)
            _sb.Clear()
        Else
            _sb.Append(val)
            If _sb.Length > 10000 Then
                If _streamWriter IsNot Nothing Then _streamWriter.Write(_sb.ToString())
                _sb.Clear()
            End If
        End If
    End Sub

    Public Sub WriteLine() Implements IWriter.WriteLine
        _sb.AppendLine()
        If _sb.Length > 10000 Then
            If _streamWriter IsNot Nothing Then _streamWriter.Write(_sb.ToString())
            _sb.Clear()
        End If
    End Sub

    Public Sub WriteString(val As String) Implements IWriter.WriteString
        WriteJSonString(_sb, val)
    End Sub

    Public Sub Flush() Implements IWriter.Flush
        If _streamWriter IsNot Nothing Then
            _streamWriter.Write(_sb.ToString())
            _streamWriter.Flush()
        End If
        _sb.Clear()
    End Sub
End Class

Public Class InMemoryWriter
    Implements IWriter

    Private _sb As New Text.StringBuilder()

    Public Property Stream As IO.Stream Implements IWriter.Stream

    Public Sub New(stream As IO.Stream)
        If stream IsNot Nothing Then _Stream = stream
    End Sub

    Public Sub Write(val As String) Implements IWriter.Write
        _sb.Append(val)
    End Sub

    Public Sub WriteLine() Implements IWriter.WriteLine
        _sb.AppendLine()
    End Sub

    Public Sub WriteString(val As String) Implements IWriter.WriteString
        WriteJSonString(_sb, val)
    End Sub

    Public Sub Flush() Implements IWriter.Flush
        If _Stream IsNot Nothing Then
            Dim sw As New IO.StreamWriter(_Stream, System.Text.Encoding.UTF8)
            sw.Write(_sb.ToString())
            _sb.Clear()
            sw.Close()
        End If
    End Sub

End Class