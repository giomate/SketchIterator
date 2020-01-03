Imports Inventor
Imports System
Public Class Optimizador
    Dim variables(), sorted() As DimDescriptor
    Dim cantidad, repetidos, indexPrio As Integer
    Dim priorities() As Double
    Public ErrorOptimizer() As Double
    Dim delta, gain, obj, maxRes, minRes, precision As Double
    Public sp, resolution As Double
    Public Sub New(ByRef optVariables() As DimDescriptor)
        variables = optVariables
        cantidad = variables.Length
        repetidos = 0
        resolution = 1000
        minRes = 100
        maxRes = 100000
        ReDim ErrorOptimizer(cantidad - 1)
        SortbyPriority()
        optVariables = sorted
    End Sub
    Function HighPriority() As String
        Dim prio(cantidad - 1), value As Double
        Dim hp As String = Nothing
        Dim a, b, indices(0), i As Integer
        Try

            For Each variable As DimDescriptor In variables
                i = Array.IndexOf(variables, variable)
                prio.SetValue(variable.PO.Prio, i)
            Next
            a = Array.IndexOf(prio, prio.Min)
            b = Array.LastIndexOf(prio, prio.Min)
            If a = b Then
                hp = variables(a).PO.Name
            Else
                For Each value In prio
                    If value = prio(a) Then
                        indices(indices.Length - 1) = Array.IndexOf(prio, value)
                        ReDim Preserve indices(indices.Length)
                    End If

                Next
                For Each i In indices
                    If variables(i).PO.Type = DimDescriptor.DimensionType.angulo Then
                        prio(i) = 0
                    ElseIf (variables(i).PO.Ceil - variables(i).PO.Floor) < (variables(i + 1).PO.Ceil - variables(i + 1).PO.Floor) Then
                        prio(i + 1) = prio(i) * (1 + Math.Exp(-CDbl(Array.IndexOf(indices, i))))
                    End If


                Next
                hp = variables(Array.IndexOf(prio, prio.Min)).PO.Name

            End If
        Catch ex As Exception
            Debug.Print(ex.ToString())
        End Try
        Return hp
    End Function
    Sub SortbyPriority()

        Dim prio(cantidad - 1), toSort(cantidad - 1), value As Double
        Dim indices(0), i, j As Integer
        Dim tempVariables(cantidad - 1) As DimDescriptor

        Try
            eliminatePrioDuplicates(toSort)
            Array.Sort(toSort)
            prio = toSort
            For Each value In prio
                i = Array.IndexOf(prio, value)
                For j = 0 To cantidad - 1
                    If prio(i) = variables(j).PO.Prio Then
                        tempVariables(i) = New DimDescriptor
                        tempVariables(i) = variables(j)
                        j = cantidad
                    End If
                Next
            Next
            sorted = tempVariables


        Catch ex As Exception
            Debug.Print(ex.ToString())
        End Try




    End Sub
    Function EliminatePrioDuplicates(ByRef sorted() As Double) As Integer
        Dim prio(cantidad - 1), rango(cantidad - 1), value As Double
        Dim i, j As Integer
        For Each variable As DimDescriptor In variables
            i = Array.IndexOf(variables, variable)
            prio.SetValue(variable.PO.Prio, i)
            rango(i) = (variable.PO.Ceil - variable.PO.Floor)

        Next
        If Not prio.Length = prio.Distinct.ToArray.Length Then
            For Each value In prio
                i = Array.IndexOf(prio, value)
                For j = i + 1 To prio.Distinct.ToArray.Length
                    If value = prio(j) Then
                        prio(i) = prio(i) * (1 + Math.Exp(-rango(j) / rango(i)))
                        variables(i).PO.Prio = prio(i)
                    End If
                Next

            Next
        End If
        sorted = prio

        Return prio.Length
    End Function
    Function Ganancia(setValue As Double, current As Double) As Double
        Try

            delta = (setValue - current) / (setValue * resolution)
            gain = Math.Exp(delta)


        Catch ex As Exception
            Debug.Print(ex.ToString())
            Debug.Print("Fail Calculating Ganacia")
        End Try

        Return gain
    End Function
    Function Ganancia(setValue As Double, current As Double, counter As Integer) As Double
        Dim g, e As Double
        g = Ganancia(setValue, current)
        e = CalculateError(counter)
        Return g
    End Function
    Function CalculateError(index As Integer) As Double
        Dim e As Double

        e = delta * resolution
        ErrorOptimizer(index) = e

        Return e
    End Function
    Function ObjetiveZeroError() As Double
        Dim obj As Double = 0
        Dim i As Integer
        For Each variable As DimDescriptor In variables
            i = Array.IndexOf(variables, variable)
            obj = obj + SingleError(i)

        Next


        Return obj
    End Function
    Function SingleError(i As Integer) As Double
        Dim obj As Double = 0
        obj = Math.Pow(ErrorOptimizer(i), 2) * Math.Exp(-variables(i).PO.Prio)
        Return obj
    End Function
    Public Sub IncrementResolution()
        If resolution < maxRes Then
            resolution = resolution + 2
        End If
        Debug.Print("Resolution:  " & resolution.ToString)
    End Sub
    Public Sub DecrementResolution()
        If resolution > minRes Then
            resolution = resolution - 1
        End If
        Debug.Print("Resolution:  " & resolution.ToString)
    End Sub
    Public Function IsPrecise(setpoint As Double, current As Double) As Boolean
        Dim b As Boolean
        precision = Math.Abs(GetDelta(setpoint, current) * resolution)

        b = precision > (setpoint / resolution)
        Return Not b
    End Function
    Public Function GetDelta(setValue As Double, current As Double) As Double
        delta = (setValue - current) / (setValue * resolution)
        Return delta
    End Function
    Public Function GetDelta() As Double

        Return delta
    End Function



End Class
