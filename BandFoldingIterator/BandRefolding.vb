Imports Inventor
Imports BandFoldingIterator.Iterator


Public Class BandRefolding
    Dim Sk3D As Sketch3D
    Dim oDoc As PartDocument
    Dim alpha, beta, setPoint, ladoa, ladob, ladoc As Double

    Dim comando As Comandos
    Dim sketchos() As SketchOptimizer
    Dim label As New Nombres
    Dim cantidadSketches As Integer
    Public healthy, running, done, driftDone, revision As Boolean
    Dim monitor As DesignMonitoring
    Dim optimoBanda As Optimizador
    Dim medicoBanda As DesignDoctor
    Public Structure ParametersCollection
        Public a As Parameter
        Public b As Parameter
        Public c As Parameter
    End Structure
    Public lados As ParametersCollection


    Public Sub New(docu As Inventor.Document)
        oDoc = docu
        comando = New Comandos(docu.Parent)
        monitor = New DesignMonitoring(oDoc)
        GetSketches()
        optimoBanda = New Optimizador(FindMainSketcho().optVariables)
        healthy = True
        driftDone = False
        optimoBanda.DownScale(1)
        medicoBanda = New DesignDoctor(docu)

    End Sub
    Public Sub CloseDocument()
        oDoc.Close()

    End Sub
    Function SetApha(a As Double) As Double
        alpha = a
        Return a
    End Function
    Function GetBeta() As Double
        beta = GetParameter("beta")._Value
        Return beta
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
    Function EstimateSetPoint(a As Double) As Double
        Dim d, p As Double

        p = FindMainSketcho().GetParameter(optimoBanda.HighPriority())._Value
        d = optimoBanda.Ganancia(a, p) * p
        Return d
    End Function
    Function StartFolding(a As Double) As Boolean
        running = True
        done = False
        Try
            alpha = a
            While (Not IsAcomplish(a) And healthy)
                running = MakeSmallDrift(a)
            End While

            beta = GetBeta()

        Catch ex As Exception
            healthy = False
            Debug.Print(ex.ToString())
            RecoverDoc(FindMainSketch.Name)

        End Try

        running = False
        Return done
    End Function
    Function StartFoldingAutomatic() As Boolean
        running = True
        done = False

        Try
            oDoc.Activate()

            While (Not IsAcomplish(lados.a._Value) And healthy)
                running = MakeSmallDriftByIndex(lados.a._Value, 0)
                While (Not IsAcomplish(lados.b._Value) And healthy)
                    running = MakeSmallDriftByIndex(lados.b._Value, 1)
                    While (Not IsAcomplish(lados.c._Value) And healthy)
                        running = MakeSmallDriftByIndex(lados.c._Value, 2)
                    End While
                End While
            End While

            ' beta = GetBeta()

        Catch ex As Exception
            healthy = False
            Debug.Print(ex.ToString())
            RecoverDoc(FindMainSketch.Name)

        End Try

        running = False
        Return done
    End Function
    Function MakeSmallDrift(a As Double) As Boolean
        Dim s As Double

        While (IsDocHealthy() And (Not IsDriftDone()))
            s = EstimateSetPoint(a)
            If ChangeMainObjetiveByIndex(s * 10000, 0) = s * 10000 Then

                StartSequence(FindMainSketch().Name)
            End If
        End While
        Return IsDriftDone()
    End Function
    Function MakeSmallDriftByIndex(a As Double, i As Integer) As Boolean
        Dim s As Double

        While (IsDocHealthy() And (Not IsDriftDone()))
            s = EstimateSetPoint(a)
            If ChangeMainObjetiveByIndex(s * 10000, i) = s * 10000 Then

                StartSequence(FindMainSketch().Name)
            End If
        End While
        Return IsDriftDone()
    End Function
    Function IsDriftDone() As Boolean
        Return FindMainSketcho.GotTarget(optimoBanda.HighPriority())
    End Function

    Function IsDocHealthy() As Boolean
        Try
            If monitor.PartHasProblems(oDoc) Then
                healthy = False

            Else
                If oDoc._SickNodesCount > 0 Then
                    healthy = False
                    medicoBanda.UndoFeature(monitor.sickFeature)
                    IsDocHealthy()
                Else
                    healthy = True
                End If
            End If
        Catch ex As Exception
            Debug.Print(ex.ToString())
            healthy = False
        End Try


        Return healthy
    End Function
    Public Function Run() As Boolean
        running = True
        done = False
        Try
            StartSequence(FindMainSketch().Name)
            If done Then
                Debug.Print("!!! done !!!")
            End If
            running = False
        Catch ex As Exception
            done = False
            Debug.Print(ex.ToString())
            running = False
        End Try
        Return done
    End Function
    Function FindMainSketch() As Sketch3D
        Return FindSketch("s0")
    End Function
    Function FindSketch(s As String) As Sketch3D
        For Each sketch As SketchOptimizer In sketchos
            If sketch.Sk3D.Name = s Then
                Return sketch.Sk3D
            End If
        Next
        Return sketchos(0).Sk3D
    End Function
    Function GetSketches() As Integer
        Dim i As Integer = 0
        For Each sketch As Sketch3D In oDoc.ComponentDefinition.Sketches3D
            If label.ContainSFirst(sketch.Name) Then
                ReDim Preserve sketchos(i)
                sketchos(i) = New SketchOptimizer(sketch.Name, oDoc)
                i = i + 1
            End If
        Next
        cantidadSketches = i
        Return i
    End Function
    Function GetInitialParametersSketch() As Sketch3D
        For Each sketch As Sketch3D In oDoc.ComponentDefinition.Sketches3D
            If sketch.Name = "sdf" Then
                Return sketch
            End If
        Next

        Return Nothing
    End Function
    Function IsOptSketch(n As String) As Boolean
        Dim b As Boolean
        If label.ContainSFirst(n) Then
            b = True
        End If
        Return b
    End Function
    Function IsAcomplish(a As Double) As Boolean

        Return optimoBanda.IsPrecise(a, GetMainVariable())
    End Function
    Function GetMainVariable() As Double
        Dim d As Double
        d = FindMainSketcho().GetParameter(optimoBanda.HighPriority())._Value
        Return d
    End Function
    Function ChangeMainObjetiveByIndex(a As Double, i As Integer) As Double
        Dim sketcho As SketchOptimizer = FindMainSketcho()
        If Not sketcho.optVariables(i).PO.Setpoint = a Then
            sketcho.optVariables(i).PO.Setpoint = a
        End If
        Return sketcho.optVariables(i).PO.Setpoint
    End Function

    Function FindMainSketcho() As SketchOptimizer
        For Each sketcho As SketchOptimizer In sketchos
            If FindMainSketch().Equals(sketcho.Sk3D) Then
                Return sketcho
            End If
        Next
        Return sketchos(0)
    End Function
    Function FindSketcho(s As String) As SketchOptimizer
        For Each sketcho As SketchOptimizer In sketchos
            If sketcho.Sk3D.Name = s Then
                Return sketcho
            End If
        Next
        Return sketchos(0)
    End Function


    Function StartSequence(s As String) As Integer
        Dim i, j, k As Integer

        Try
            i = Array.IndexOf(sketchos, FindSketcho(s))
            j = sketchos.Length - 1
            k = j
            While IsDocHealthy() And ((j >= i) And (j <= k)) And ((Not IsDriftDone()) Or revision)
                driftDone = False
                sketchos(j).Run()
                If sketchos(j).done Then
                    sketchos(j).done = False
                    If j < k And j >= i Then
                        revision = True
                        StartSequence(sketchos(j + 1).Sk3D.Name)
                        revision = False

                    End If
                    j = j - 1

                Else
                    If sketchos(j).sick Then
                        sketchos(j).sick = False
                        If j = i Then
                            Debug.Print("Not posible")
                            StartAgain()
                        Else
                            j = j + 1
                        End If


                    End If
                End If
            End While


        Catch ex As Exception
            Debug.Print(ex.ToString())
            StartOver(s)
        End Try

        Return j
    End Function
    Function IsFoldingDone(s As String) As Boolean
        Dim i As Integer
        i = Array.IndexOf(sketchos, FindSketcho(s))

        IsFoldingDone = IsSketchInRange(s)
        If IsFoldingDone Then
            If s = FindMainSketch().Name Then
                IsFoldingDone = IsAcomplish(alpha)
                If IsFoldingDone Then
                    driftDone = True
                    done = True
                    Return True
                Else
                    Return False
                End If
            Else
                Return True
            End If

        End If

        Return False
    End Function
    Function IsSketchInRange(s As String) As Boolean
        Dim i As Integer
        i = Array.IndexOf(sketchos, FindSketcho(s))
        Return sketchos(i).AreAllInRange()
    End Function

    Function StartOver(s As String) As Boolean
        Dim i As Integer
        Dim b As Boolean = False
        Try
            i = Array.IndexOf(sketchos, FindSketcho(s))
            RecoverDoc(s)
            If i > sketchos.Length - 1 Then
                StartAgain()
            Else
                b = StartSequence(sketchos(i + 1).Sk3D.Name)
            End If
        Catch ex As Exception
            StartAgain()
            Debug.Print("StartAgain")
        End Try
        Return b

    End Function

    Function RecoverDoc(s As String) As Boolean
        Dim ok As Boolean = False
        Try
            While ((Not ok) And comando.IsUndoable())
                comando.UndoCommand()
                Debug.Print("undoing")
                ok = IsFeasible(s)

            End While

        Catch ex As Exception
            Debug.Print(ex.ToString())
            StartOver(s)
        End Try
        Return ok
    End Function
    Public Function IsFeasible(s As String) As Boolean
        Dim b As Boolean = False
        Try
            oDoc.Update2(True)
            If IsDocHealthy() Then
                b = True
                Debug.Print("!!Doc Healthy!!")
            Else
                b = False
            End If
            Return b
        Catch ex As Exception
            Debug.Print(ex.ToString())

            StartOver(s)
            Return b

        End Try

    End Function
    Function StartAgain() As Boolean
        Dim b As Boolean = False
        Try
            RecoverDoc(FindMainSketch.Name)
            Debug.Print("StartAgain")
            b = Run()

        Catch ex As Exception
            b = False
            Debug.Print("StartAgainAgain")

            Return b
        End Try
        Return b

    End Function

End Class
