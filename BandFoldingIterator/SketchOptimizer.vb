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
    Public healthy, running, done, sick As Boolean
    Dim monitor As DesignMonitoring
    Dim medico As DesignDoctor



    Public Sub New(sketchName As String, docu As Inventor.Document)
        oDoc = docu
        Sk3D = oDoc.ComponentDefinition.Sketches3D.Item(sketchName)
        comando = New Comandos(docu.Parent)
        traductor = New Parser(Sk3D)
        traductor.FillValues(optVariables)
        optimo = New Optimizador(optVariables)
        monitor = New DesignMonitoring(docu)
        medico = New DesignDoctor(docu)
        sick = False
        done = False
    End Sub
    Public Sub OpenSketch()
        Sk3D = oDoc.ComponentDefinition.Sketches3D.Item(sketchName)
    End Sub
    Public Sub OpenSketch(sketchName As String)
        Sk3D = oDoc.ComponentDefinition.Sketches3D.Item(sketchName)
    End Sub
    Function Run() As Boolean
        running = True
        done = False
        Try
            If Not sick Then
                If traductor.IsVariableInSketch(optimo.HighPriority()) Then
                    MakeSketchesInvisible()
                    Sk3D.Visible = True
                    If adjustParameter(optimo.HighPriority()) Then
                        Debug.Print("!!! done !!!")
                        MakeallDriven()
                        If IsBuilt(optimo.HighPriority()) Then
                            If IsStatusDocOk(oDoc) Then
                                done = True
                                MakeallDriven()
                            Else
                                StartOver(optimo.HighPriority())
                            End If

                        End If

                    End If
                    MakeSketchesInvisible()
                End If


            Else
                done = False
            End If


            running = False
        Catch ex As Exception
            Debug.Print(ex.ToString())
            running = False
        End Try
        Return done
    End Function


    Function adjustParameter(name As String) As Boolean
        Dim p As Parameter = Nothing

        Try



            'Sk3D.Edit()
            MakeallDriven()
            CheckOtherVariables(name)
            If MainIteration(name) > 0 Then
                If Sk3D.Visible Then
                    Sk3D.Solve()
                    'Sk3D.ExitEdit()
                    If IsBuilt(name) Then
                        If IsStatusDocOk(oDoc) Then
                            sick = False

                        Else
                            StartOver(name)
                        End If

                    End If
                End If


            Else
                RecoverDocument(name)
            End If




        Catch ex4 As Exception
            Sk3D.Visible = True

            p = GetParameter(name)
            Debug.Print(ex4.ToString())
            Debug.Print("Fail adjusting " & name & " ...last value:" & p.Value.ToString)
            StartOver(name)
            Return False
        End Try
        Return Not sick
    End Function
    Function MakeSketchesInvisible() As Integer

        Dim i As Integer = 0
        For Each sketch As Sketch3D In oDoc.ComponentDefinition.Sketches3D
            sketch.Visible = False
            i = i + 1
        Next

        Return i

    End Function
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
                    Return 1
                Else
                    While ((Not GotTarget(name)) And (Not sick))
                        p = Iterate(name)
                        CheckOtherVariables(name)
                        CalculateGain(name)
                    End While
                End If
                If sick Then
                    Return 0
                End If
            End If

        Catch ex As Exception


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

        Dim g As Double

        Try
            If name = optimo.HighPriority Then
                MakeallDrivenExc(name)
            Else
                GetDimension(name).Driven = False
            End If

            pit = GetParameter(name)
            If IsSolvable(name) Then
                While ((Not IsPreciso(name)) And (Not sick))
                    g = CalculateGain(name)
                    pit.Value = pit.Value * g
                    Debug.Print("iterating  " & pit.Name & " = " & pit.Value.ToString)

                    If IsFoldable(name) Then
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
                While ((Not IsBetweenRange(name)) And (Not sick))
                    g = CalculateGain(name)
                    pit.Value = pit.Value * g
                    Debug.Print(" Out of the Range  " & pit.Name & " = " & pit.Value.ToString)

                    If IsFoldable(name) Then
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


            healthy = False
            Debug.Print(ex.ToString())
            Debug.Print("Fail adjusting " & name & " ...last value:" & pit.Value.ToString)
            StartOver(name)
            Return pit

        End Try

        Return pit
    End Function
    Sub StartOver(name As String)
        Try
            RecoverSketch(name)
            optimo.IncrementResolution()
            If Not RecoverSketch(name) Then
                MakeallDriven()
                RecoverSketch(name)
            Else
                If GetIndexVariable(name) < optVariables.Length - 1 Then
                    adjustParameter(optVariables(GetIndexVariable(name) + 1).PO.Name)
                Else
                    Debug.Print("StartOver")
                    StartAgain(name)
                End If
            End If


        Catch ex As Exception
            Debug.Print("StartOver")
            StartAgain(name)
        End Try


    End Sub
    Sub StartAgain(name As String)
        RecoverSketch(name)
        optimo.IncrementResolution()
        MakeallDriven()
        If GetIndexVariable(name) < optVariables.Length - 1 Then
            adjustParameter(optVariables(GetIndexVariable(name) + 1).PO.Name)
        Else

            Debug.Print("StartAgain")
            sick = True
            Run()
        End If

    End Sub
    Public Function IsEditable(n As String) As Boolean
        Dim h As Boolean = True
        Try
            If GetDimension(n).Driven Then
                GetDimension(n).Driven = False
            End If


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
                If monitor.AreDimensionsHealthy(Sk3D) Then
                    Sk3D.Solve()
                Else
                    MakeallDriven()
                End If

            End If

            If HealthSketch() Then
                s = True
                Debug.Print("!!Sketch Healthy!!")
            Else
                RecoverSketch(n)
            End If

        Catch ex As Exception
            Sk3D.Visible = True
            s = False
            kapput = GetParameter(n)
            Debug.Print(ex.ToString())
            Debug.Print("Fail adjusting " & n & " ...last value:" & kapput.Value.ToString)
            Debug.Print(Sk3D.HealthStatus.ToString)


            StartOver(kapput.Name)
            Return s

        End Try
        Return s
    End Function
    Public Function IsFoldable(n As String) As Boolean
        Dim s As Boolean = False
        Try
            If GetDimension(n).Driven Then
                GetDimension(n).Driven = False
            End If
            If IsSolvable(n) Then
                oDoc.Update2(True)
                If IsStatusDocOk(oDoc) Then
                    s = True
                    Debug.Print("!!Update passed!!")
                Else

                    Debug.Print("###  Feature Sick   ###" & monitor.sickFeature.ToString)
                    RecoverDocument(n)


                End If

            End If



        Catch ex As Exception
            s = False
            kapput = GetParameter(n)
            Debug.Print(ex.ToString())
            Debug.Print("Fail adjusting " & n & " ...last value:" & kapput.Value.ToString)
            Debug.Print(Sk3D.HealthStatus.ToString)

            StartOver(kapput.Name)
            Return s

        End Try
        Return s
    End Function
    Public Function RecoverSketch(name As String) As Boolean
        Dim ok As Boolean = False
        Try
            While ((Not ok) And comando.IsUndoable())
                comando.UndoCommand()
                MakeallDriven()
                Debug.Print("undoing Sketch")
                ok = IsFoldable(name)


            End While

        Catch ex As Exception
            Debug.Print(ex.ToString())
            MakeallDriven()
            StartOver(name)
        End Try
        Return ok

    End Function
    Public Function RecoverDocument(name As String) As Boolean
        Dim ok As Boolean = False
        Try
            While ((Not ok) And comando.IsUndoable())
                comando.UndoCommand()
                MakeallDriven()
                Debug.Print("undoing Document")
                ok = IsStatusDocOk(oDoc)
                If ok Then
                    Sk3D.Solve()
                    'Sk3D.ExitEdit()
                    ok = IsStatusDocOk(oDoc)
                    If Not ok Then
                        sick = True
                    End If
                End If

            End While

        Catch ex As Exception
            sick = True
            Debug.Print(ex.ToString())
            MakeallDriven()
            StartOver(name)
        End Try
        Return ok

    End Function
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
        If Sk3D.HealthStatus = HealthStatusEnum.kOutOfDateHealth Or
         Sk3D.HealthStatus = HealthStatusEnum.kUpToDateHealth Then
            If monitor.AreDimensionsHealthy(Sk3D) Then
                If monitor.AreConstrainsHealthy(Sk3D) Then
                    b = True
                Else
                    Debug.Print("not deletable constrain")
                    b = True
                End If
            Else
                b = False

            End If
        ElseIf Sk3D.HealthStatus = HealthStatusEnum.kDriverLostHealth Then
            MakeallDriven()
            b = HealthSketch()

        Else
            b = False
        End If
        Return b
    End Function
    Function IsStatusDocOk(partDoc As PartDocument) As Boolean
        Dim b As Boolean = True
        Try
            If monitor.PartHasProblems(partDoc) Then
                b = medico.UndoFeature(monitor.sickFeature)
                If b Then
                    b = Not monitor.PartHasProblems(partDoc)
                Else
                    sick = True
                End If

            End If
        Catch ex As Exception
            Debug.Print(ex.ToString())
            StartAgain(optimo.HighPriority())
        End Try



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
        Dim v As DimDescriptor
        Try
            p = GetParameter(name)
            g = optimo.Ganancia(GetSetPoint(name), p._Value, GetIndexVariable(name))
            v = optVariables(GetIndexVariable(name))
            If (v.PO.Type = DimDescriptor.DimensionType.angulo) Then
                g = Math.Sqrt(g)

            End If
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
