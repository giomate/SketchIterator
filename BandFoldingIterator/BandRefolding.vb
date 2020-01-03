Imports Inventor


Public Class BandRefolding
    Dim Sk3D As Sketch3D
    Dim oDoc As PartDocument
    Public alpha, beta As Double
    Dim comando As Comandos
    Dim sketchos() As SketchOptimizer
    Dim label As New Nombres
    Dim cantidadSketches As Integer
    Public healthy, running, done As Boolean


    Public Sub New(docu As Inventor.Document)
        oDoc = docu
        comando = New Comandos(docu.Parent)
        GetSketches()
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
    Function StartFolding(a As Double) As Boolean
        running = True
        done = False
        Try

            If ChangeMainObjetive(a * 1000) = a * 1000 Then
                While IsDocHealthy()
                    StartSequence(FindMaindSketch.Name)
                End While
            End If
            beta = GetBeta()

        Catch ex As Exception
            healthy = False
            Debug.Print(ex.ToString())
            RecoverDoc(FindMaindSketch.Name)

        End Try
        done = True
        running = False
        Return done
    End Function

    Function IsDocHealthy() As Boolean
        If oDoc._SickNodesCount > 0 Then
            healthy = False
        Else
            healthy = True
        End If
        Return healthy
    End Function
    Public Function Run() As Boolean
        running = True
        done = False
        Try
            StartSequence(FindMaindSketch().Name)
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
    Function FindMaindSketch() As Sketch3D
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
    Function IsOptSketch(n As String) As Boolean
        Dim b As Boolean
        If label.ContainSFirst(n) Then
            b = True
        End If
        Return b
    End Function
    Function ChangeMainObjetive(a As Double) As Double
        Dim sketcho As SketchOptimizer = FindMaindSketcho()
        If Not sketcho.optVariables(0).PO.Setpoint = a Then
            sketcho.optVariables(0).PO.Setpoint = a
        End If
        Return sketcho.optVariables(0).PO.Setpoint
    End Function
    Function FindMaindSketcho() As SketchOptimizer
        For Each sketcho As SketchOptimizer In sketchos
            If FindMaindSketch().Equals(sketcho.Sk3D) Then
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
    Function StartMainIterator() As Boolean
        Dim b As Boolean

        Try
            b = StartIterator(FindMaindSketch.Name)
        Catch ex As Exception

        End Try
        Return b
    End Function
    Function StartIterator(s As String) As Boolean
        Dim b As Boolean = False
        Dim i As Integer
        Try
            i = Array.IndexOf(sketchos, FindSketcho(s))
            While (IsDocHealthy() And (i >= 0))

                If Not sketchos(i).running Then
                    sketchos(i).Run()

                End If

            End While
        Catch ex As Exception
            Debug.Print(ex.ToString())
            StartOver(s)
        End Try
        Return b
    End Function
    Function StartSequence(s As String) As Integer
        Dim i, j As Integer
        Dim b As Boolean = False
        Try
            i = Array.IndexOf(sketchos, FindSketcho(s))
            j = sketchos.Length - 1
            While IsDocHealthy() And (j >= i)
                sketchos(j).Run()
                If sketchos(j).done Then
                    If j = i Then
                        b = True
                    End If
                    j = j - 1
                End If
            End While


        Catch ex As Exception
            Debug.Print(ex.ToString())
            StartOver(s)
        End Try

        Return 0
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
        Dim h As Boolean = False
        Try
            While ((Not h) And comando.IsUndoable())
                comando.UndoCommand()
                Debug.Print("undoing")
                If IsFeasible(s) Then
                    h = True
                End If

            End While
            h = IsDocHealthy()
        Catch ex As Exception
            Debug.Print(ex.ToString())
            StartOver(s)
        End Try
        Return h
    End Function
    Public Function IsFeasible(s As String) As Boolean
        Dim b As Boolean = False
        Try
            oDoc.Rebuild2(True)
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
            RecoverDoc(FindMaindSketch.Name)
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
