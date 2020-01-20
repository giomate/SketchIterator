Imports Inventor
Imports System.Text.RegularExpressions

Public Class Parser

    Dim oSk3D As Sketch3D
    Dim oPartDoc As PartDocument
    Dim sketchName, variablesNames(), codigo, resto As String
    Dim dimensions() As DimensionConstraint3D
    Dim optVariables() As DimDescriptor
    Dim indexOpt, cantidad As Integer


    Public Sub New(sketchFile As Sketch3D)
        oSk3D = sketchFile
        sketchName = oSk3D.Name
        cantidad = GetVariableSketchName()

    End Sub
    Public Sub FillValues(ByRef optMatrix() As DimDescriptor)
        If ExtractValues() > 0 Then

            optMatrix = optVariables
        End If


    End Sub

    Function GetVariableNames() As Integer
        For Each dimension In oSk3D.DimensionConstraints3D

        Next

        Return 0
    End Function

    Function GetVariableSketchName() As Integer
        Dim i As Integer = 0
        Dim pattern = String.Concat("\b", sketchName, "\w*")
        Try
            For Each dimension As DimensionConstraint3D In oSk3D.DimensionConstraints3D
                If Regex.IsMatch(dimension.Parameter.Name, pattern) Then
                    ReDim Preserve dimensions(i)
                    ReDim Preserve variablesNames(i)
                    dimensions(i) = dimension
                    variablesNames(i) = dimension.Parameter.Name
                    ' ReDim Preserve variablesNames(UBound(variablesNames) + 1)
                    ' ReDim Preserve dimensions(UBound(dimensions) + 1)
                    i = i + 1
                End If
            Next
        Catch ex As Exception
            Debug.Print(ex.ToString())
        End Try

        cantidad = dimensions.Length
        Return cantidad
    End Function
    Public Function IsVariableInSketch(variable As String) As Boolean


        Try
            For Each dimension As DimensionConstraint3D In oSk3D.DimensionConstraints3D
                If Regex.IsMatch(dimension.Parameter.Name, variable) Then
                    Return True
                End If
            Next
        Catch ex As Exception
            Debug.Print(ex.ToString())
        End Try


        Return False
    End Function
    Function ExtractValues() As Integer
        ReDim optVariables(cantidad - 1)
        Dim i As Integer
        For Each dimension As DimensionConstraint3D In dimensions
            i = Array.IndexOf(dimensions, dimension)
            indexOpt = i
            optVariables(i) = New DimDescriptor
            optVariables(i).PO.Name = dimension.Parameter.Name
            GetTypeDimension(dimension.Parameter.Name)
            GetPrio(resto)
            GetSetPoint(resto)
            GetLimits(resto)
        Next

        Return indexOpt
    End Function
    Function GetTypeDimension(dimensionName As String) As DimDescriptor.DimensionType
        Dim p() As String
        Dim tipo As DimDescriptor.DimensionType
        p = Strings.Split(dimensionName, sketchName)
        If Regex.IsMatch(p(1), "d") Then
            tipo = DimDescriptor.DimensionType.dimension
        ElseIf Regex.IsMatch(p(1), "a") Then
            tipo = DimDescriptor.DimensionType.angulo
        Else
            tipo = DimDescriptor.DimensionType.dimension

        End If
        OptVariables(indexOpt).PO.Type = tipo
        resto = p(1)
        Return tipo
    End Function
    Function GetPrio(pedazo As String) As Double
        Dim p() As String
        Dim prio As Double
        p = Strings.Split(pedazo, "p")
        resto = p(1)
        p = Strings.Split(p(1), "z")
        prio = CDbl(p(0))

        OptVariables(indexOpt).PO.Prio = prio
        Return prio
    End Function
    Function GetSetPoint(pedazo As String) As Double
        Dim p() As String
        Dim sp As Double
        p = Strings.Split(pedazo, "z")
        resto = p(1)
        p = Strings.Split(p(1), "c")
        sp = CDbl(p(0))

        OptVariables(indexOpt).PO.Setpoint = sp
        Return sp
    End Function
    Function GetLimits(pedazo As String) As Double
        Dim p() As String
        Dim c, f As Double
        p = Strings.Split(pedazo, "c")

        p = Strings.Split(p(1), "f")
        c = CDbl(p(0))
        f = CDbl(p(1))
        optVariables(indexOpt).PO.Ceil = c
        optVariables(indexOpt).PO.Floor = f
        Return c
    End Function

End Class
