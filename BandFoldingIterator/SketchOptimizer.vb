Imports Inventor
Public Class SketchOptimizer
    Public Sk3D As Sketch3D
    Dim oDoc As PartDocument
    Dim sketchName As String
    Public optVariables() As DimDescriptor
    Dim dimension As DimensionConstraint3D
    Public optimo As Optimizador
    Dim errorOpt, gain As Double
    Dim traductor As Parser
    Dim comando As Comandos
    Dim parametro, kapput As Parameter
    Public healthy, running, done As Boolean


    Public Sub New(sketchName As String, docu As Inventor.Document)
        oDoc = docu
        Sk3D = oDoc.ComponentDefinition.Sketches3D.Item(sketchName)
        comando = New Comandos(docu.Parent)
        traductor = New Parser(Sk3D)
        traductor.FillValues(optVariables)
        optimo = New Optimizador(optVariables)
    End Sub
    Public Sub OpenSketch()
        Sk3D = oDoc.ComponentDefinition.Sketches3D.Item(sketchName)
    End Sub
    Public Sub OpenSketch(sketchName As String)
        Sk3D = oDoc.ComponentDefinition.Sketches3D.Item(sketchName)
    End Sub
    Function Run() As Double
        running = True
        done = False
        Try

            adjustParameter(optimo.HighPriority())
            If done Then
                Debug.Print("!!! done !!!")
            End If
            running = False
        Catch ex As Exception
            Debug.Print(ex.ToString())
            running = False
        End Try
        Return CurrentError(optimo.HighPriority())
    End Function

    Public Sub adjustParameter(name As String)
        Dim p As Parameter = Nothing

        Try

            p = GetParameter(name)

            Sk3D.Edit()
            MakeallDriven()
            CheckOtherVariables(name)
            If MainIteration(name) > 0 Then
                Sk3D.Solve()
                Sk3D.ExitEdit()
                If IsBuilt(name) Then
                    done = True
                End If
            End If




        Catch ex4 As Exception

            comando.UndoCommand()
            MakeallDriven()
            optimo.IncrementResolution()

            Debug.Print(ex4.ToString())
            Debug.Print("Fail adjusting " & name & " ...last value:" & p.Value.ToString)
            StartOver(name)
            Exit Sub
        End Try


    End Sub
    Function IsBuilt(name As String) As Boolean
        Dim b As Boolean
        Try
            If oDoc.Dirty Then
                If GetDimension(name).Driven Then
                    GetDimension(name).Driven = False
                End If
                oDoc.Update2(True)
                If oDoc._SickNodesCount > 0 Then
                    StartOver(name)
                Else
                    b = True
                End If
            Else
                b = True
            End If

        Catch ex As Exception
            b = False
            Debug.Print(ex.ToString())
            StartOver(name)
            Return b
        End Try
        Return b
    End Function
    Function MainIteration(name As String) As Double
        Dim p As Parameter = Nothing
        Try

            MakeallDrivenExc(name)
            If IsSolvable(name) Then
                If GotTarget(name) Then
                    done = True
                Else
                    While Not (GotTarget(name))
                        p = Iterate(name)
                        CheckOtherVariables(name)
                        CalculateGain(name)
                    End While
                End If
            End If

        Catch ex As Exception
            comando.UndoCommand()
            MakeallDriven()
            optimo.IncrementResolution()

            Debug.Print(ex.ToString())
            Debug.Print("Fail adjusting " & p.Name & " ...last value:" & p.Value.ToString)
            StartOver(p.Name)
            Return -1
        End Try

        Return optimo.ObjetiveZeroError()
    End Function
    Function GotTarget(name As String) As Boolean
        Dim b As Boolean
        Dim p As Parameter = Nothing
        Dim sp, meta As Double
        If name = optimo.HighPriority() Then
            meta = optimo.ObjetiveZeroError()
        Else
            meta = optimo.SingleError(GetIndexVariable(name))
        End If
        b = ((meta < sp / optimo.resolution) Or IsPreciso(name))

        Return b
    End Function
    Public Function Iterate(name As String) As Parameter
        Dim pit As Parameter = Nothing

        Dim g, sp As Double

        Try

            GetDimension(name).Driven = False
            sp = GetSetPoint(name)
            pit = GetParameter(name)
            If IsSolvable(name) Then
                While (Not IsPreciso(name))
                    g = CalculateGain(name)
                    pit.Value = pit.Value * g
                    Debug.Print("iterating  " & pit.Name & " = " & pit.Value.ToString)

                    If IsSolvable(name) Then
                        CheckOtherVariables(pit.Name)
                        optimo.DecrementResolution()
                        healthy = True
                        Debug.Print("Sketch Solved")
                    End If

                End While
                GetDimension(pit.Name).Driven = True
            Else
                StartOver(pit.Name)
            End If



        Catch ex As Exception
            Debug.Print(ex.ToString())
            Debug.Print("Fail adjusting " & name & " ...last value:" & pit.Value.ToString)
            RecoverSketch(name)
            optimo.IncrementResolution()
            healthy = False

            StartOver(name)
            Return pit

        End Try

        Return pit
    End Function
    Public Function IterateOutRange(name As String) As Parameter
        Dim pit As Parameter = Nothing

        Dim g, sp As Double

        Try

            GetDimension(name).Driven = False
            sp = GetSetPoint(name)
            pit = GetParameter(name)
            If IsSolvable(name) Then
                While (Not IsBetweenRange(name))
                    g = CalculateGain(name)
                    pit.Value = pit.Value * g
                    Debug.Print(" Out of the Range  " & pit.Name & " = " & pit.Value.ToString)

                    If IsSolvable(name) Then
                        'CheckOtherVariables(pit.Name)
                        optimo.DecrementResolution()
                        healthy = True
                        Debug.Print("Sketch Solved")
                    End If

                End While
                GetDimension(pit.Name).Driven = True
            Else
                StartOver(pit.Name)
            End If



        Catch ex As Exception

            RecoverSketch(name)
            optimo.IncrementResolution()
            healthy = False
            Debug.Print(ex.ToString())
            Debug.Print("Fail adjusting " & name & " ...last value:" & pit.Value.ToString)
            StartOver(name)
            Return pit

        End Try

        Return pit
    End Function
    Sub StartOver(name As String)
        MakeallDriven()
        If GetIndexVariable(name) < optVariables.Length - 1 Then
            adjustParameter(optVariables(GetIndexVariable(name) + 1).PO.Name)
        Else

            Debug.Print("StartOver")
            adjustParameter(optVariables(0).PO.Name)
        End If

    End Sub
    Public Function IsEditable(n As String) As Boolean
        Dim h As Boolean = True
        Try
            If GetDimension(n).Driven Then
                GetDimension(n).Driven = False
            End If
            Sk3D.Edit()

            If HealthSketch() Then
                h = True
                Debug.Print("!!Sketch Healthy!!")
            Else
                RecoverSketch(n)
            End If

        Catch ex As Exception
            h = False
            kapput = GetParameter(n)
            Debug.Print(ex.ToString())
            Debug.Print("Fail adjusting " & n & " ...last value:" & kapput.Value.ToString)
            Debug.Print(Sk3D.HealthStatus.ToString)

            RecoverSketch(n)
            optimo.IncrementResolution()

            StartOver(kapput.Name)


        End Try
        Return h
    End Function
    Public Function IsSolvable(n As String) As Boolean
        Dim s As Boolean = False
        Try
            If GetDimension(n).Driven Then
                GetDimension(n).Driven = False
            End If
            If IsEditable(n) Then
                Sk3D.Solve()
            End If

            If HealthSketch() Then
                s = True
                Debug.Print("!!Sketch Healthy!!")
            Else
                RecoverSketch(n)
            End If

        Catch ex As Exception
            s = False
            kapput = GetParameter(n)
            Debug.Print(ex.ToString())
            Debug.Print("Fail adjusting " & n & " ...last value:" & kapput.Value.ToString)
            Debug.Print(Sk3D.HealthStatus.ToString)

            RecoverSketch(n)
            optimo.IncrementResolution()

            StartOver(kapput.Name)
            Return s

        End Try
        Return s
    End Function
    Sub RecoverSketch(name As String)
        Dim out As Boolean = True
        Try
            While (out And comando.IsUndoable())
                comando.UndoCommand()
                Debug.Print("undoing")

                out = Not AreOthersInRange(name)
                If IsSolvable(name) Then
                    If out Then
                        out = False
                    End If

                End If

            End While
        Catch ex As Exception
            Debug.Print(ex.ToString())
            StartOver(name)
        End Try


    End Sub
    Function AreAllInRange() As Boolean
        Dim b As Boolean = True
        For Each variable As DimDescriptor In optVariables

            b = IsBetweenRange(variable.PO.Name) And b

        Next
        Return b
    End Function
    Function AreOthersInRange(name As String) As Boolean
        Dim b As Boolean = True
        For Each variable As DimDescriptor In optVariables
            If Not variable.PO.Name = name Then
                b = IsBetweenRange(variable.PO.Name) And b
            End If
        Next
        Return b
    End Function

    Function HealthSketch() As Boolean
        Dim b As Boolean
        If Sk3D.HealthStatus = HealthStatusEnum.kOutOfDateHealth Then
            b = True
        ElseIf Sk3D.HealthStatus = HealthStatusEnum.kUpToDateHealth Then

            b = True
        ElseIf Sk3D.HealthStatus = HealthStatusEnum.kDriverLostHealth Then
            MakeallDriven()
            b = HealthSketch()

        Else
            b = False
        End If
        Return b
    End Function
    Public Sub MakeallDriven()
        For Each variable As DimDescriptor In optVariables
            GetDimension(variable.PO.Name).Driven = True
        Next

    End Sub
    Public Sub MakeallDrivenExc(name As String)
        MakeallDriven()
        GetDimension(optVariables(GetIndexVariable(name)).PO.Name).Driven = False
    End Sub
    Public Sub MakeallDrivenExc(name As String, name2 As String)
        MakeallDrivenExc(name)
        GetDimension(optVariables(GetIndexVariable(name)).PO.Name).Driven = False
    End Sub
    Public Sub MakeallConstrained()
        For Each variable As DimDescriptor In optVariables
            GetDimension(variable.PO.Name).Driven = False
        Next

    End Sub
    Public Sub MakeallConstrainedExc(name As String)
        MakeallConstrained()
        GetDimension(optVariables(GetIndexVariable(name)).PO.Name).Driven = True
    End Sub
    Public Sub CheckOtherVariables(name As String)
        Dim i, j As Integer

        For Each variable As DimDescriptor In optVariables
            If variable.PO.Name = name Then
                i = Array.IndexOf(optVariables, variable)
                For j = i To optVariables.Length - 1
                    If optVariables(j).PO.Prio > variable.PO.Prio Then

                        If Not IsBetweenRange(optVariables(j).PO.Name) Then
                            IterateOutRange(optVariables(j).PO.Name)
                        End If
                    End If

                Next


            End If
        Next


    End Sub
    Function IsBetweenRange(name As String) As Boolean
        Dim v As DimDescriptor
        Dim b As Boolean = True
        Try
            v = optVariables(GetIndexVariable(name))
            If (v.PO.Type = DimDescriptor.DimensionType.angulo) Then
                b = IsBetweenLimits(v, 1000)
            ElseIf v.PO.Type = DimDescriptor.DimensionType.dimension Then
                b = IsBetweenLimits(v, 10000)
            End If

        Catch ex As Exception
            Debug.Print(ex.ToString())
        End Try



        Return b
    End Function

    Public Function IsBetweenLimits(DD As DimDescriptor, scale As Double) As Boolean
        Dim p As Parameter
        Dim b As Boolean
        p = GetParameter(DD.PO.Name)

        If (p._Value < DD.PO.Floor / scale Or p._Value > DD.PO.Ceil / scale) Then

            b = False

        Else
            b = True
        End If

        Return b
    End Function

    Public Function GetParameter(name As String) As Parameter
        Dim p As Parameter = Nothing
        Try
            p = oDoc.ComponentDefinition.Parameters.ModelParameters.Item(name)
        Catch ex As Exception
            Try
                p = oDoc.ComponentDefinition.Parameters.ReferenceParameters.Item(name)
            Catch ex1 As Exception
                Try
                    p = oDoc.ComponentDefinition.Parameters.UserParameters.Item(name)
                Catch ex2 As Exception
                    Debug.Print(ex2.ToString())
                    Debug.Print("Parameter not found: " & name)
                End Try

            End Try

        End Try

        Return p
    End Function
    Function GetDimension(name As String) As DimensionConstraint3D
        For Each dimension In Sk3D.DimensionConstraints3D
            If dimension.Parameter.Name = name Then
                Return dimension
            End If
        Next
        Return Nothing
    End Function
    Function GetSetPoint(name As String) As Double
        Dim sp As Double
        sp = optVariables(GetIndexVariable(name)).PO.Setpoint
        Dim v As DimDescriptor
        v = optVariables(GetIndexVariable(name))
        If (v.PO.Type = DimDescriptor.DimensionType.angulo) Then
            sp = sp / 1000
        ElseIf v.PO.Type = DimDescriptor.DimensionType.dimension Then
            sp = sp / 10000
        End If

        Return sp
    End Function

    Function CalculateGain(name As String) As Double
        Dim p As Parameter = Nothing
        Dim g As Double
        Try
            p = GetParameter(name)
            g = optimo.Ganancia(GetSetPoint(name), p._Value, GetIndexVariable(name))

        Catch ex As Exception
            Debug.Print(ex.ToString())
            Debug.Print("Fail Calculating " & p.Name)
        End Try

        Return g
    End Function
    Public Function GetIndexVariable(name As String) As Integer
        Dim i As Integer

        For Each variable As DimDescriptor In optVariables
            If variable.PO.Name = name Then
                i = Array.IndexOf(optVariables, variable)
                Return i


            End If
        Next

        Return i

    End Function
    Public Function IsPreciso(name As String) As Boolean
        Dim p As Parameter
        Dim b As Boolean
        p = GetParameter(name)
        If name = optimo.HighPriority Then
            b = optimo.IsPrecise(GetSetPoint(name), p._Value)
        Else
            b = IsBetweenRange(name)
        End If
        Return b
    End Function
    Function CurrentError(name As String) As Double

        Return optimo.ErrorOptimizer(GetIndexVariable(name))
    End Function


End Class
