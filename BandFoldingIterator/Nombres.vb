Imports System
Imports System.Text.RegularExpressions
Public Class Nombres
    Dim newName, currentName, oldName, datei As String

    Public Sub New()
        newName = "noname"
        datei = ".ipt"
    End Sub

    Public Function IncrementLabel(name As String, pattern As String) As String
        Dim p() As String

        Try
            p = Strings.Split(name, pattern)
            newName = String.Concat(p(0), pattern, CStr(CInt(p(1)) + 1))
        Catch ex As Exception
            Debug.Print(ex.ToString())
            Return name
        End Try

        Return newName
    End Function
    Public Function IncrementLabelIpt(name As String, pattern As String) As String
        Dim p() As String

        Try
            p = Strings.Split(name, datei)
            newName = String.Concat(IncrementLabel(p(0), pattern), datei)

        Catch ex As Exception
            Debug.Print(ex.ToString())
            Return name
        End Try

        Return newName
    End Function
    Public Function ContainSFirst(s As String) As Boolean
        Dim b As Boolean
        Dim pattern = String.Concat("\b", "s", "\d+", "\w*")
        b = Regex.IsMatch(s, pattern)
        Return b
    End Function
End Class
