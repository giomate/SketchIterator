Imports Inventor

Public Class RebuilderClass

    Dim oApp As Inventor.Application
    Dim oDesignProjectMgr As DesignProjectManager
    Dim oPartDoc As PartDocument
    Dim oSk3D As Sketch3D
    Dim delta, gain, sp As Double
    Dim oTheta, oGap, oTecho, p As Parameter
    Dim resolution As Integer = 100
    Dim dimension As DimensionConstraint3D
    Dim kapput As Boolean


    Public Structure DesignParam
        Public p As Integer
        Public q As Integer
        Public b As Integer
        Public Dmax As Double
        Public Dmin As Double

    End Structure

    Dim DP As DesignParam
    Dim Tr As Double
    Dim Cr As Double
    Public Sub New(App As Inventor.Application)
        oApp = App
        oDesignProjectMgr = oApp.DesignProjectManager
        oPartDoc = oApp.ActiveDocument
        DP.Dmax = 200
        DP.Dmin = 32
        Tr = (DP.Dmax + DP.Dmin) / 4
        Cr = (DP.Dmax - DP.Dmin) / 4
        DP.p = 17
        DP.q = 37
    End Sub

    Function openFile(fileName As String) As PartDocument
        Try
            If oApp.Documents.Count > 0 Then
                If Not (oApp.ActiveDocument.FullFileName = createFileName(fileName)) Then
                    oPartDoc = oApp.Documents.Open(fileName, True)
                End If
            Else
                oPartDoc = oApp.Documents.Open(fileName, True)
            End If

            oPartDoc = oApp.ActiveDocument
        Catch ex3 As Exception
            MsgBox(ex3.ToString())
            MsgBox("Unable to find Document")
        End Try

        ' Conversions.SetUnitsToMetric(oPartDoc)
        Return oPartDoc
    End Function

    Function createFileName(fileName As String) As String
        Dim strFilePath As String
        strFilePath = oDesignProjectMgr.ActiveDesignProject.WorkspacePath
        ' Dim strFileName As String
        'strFileName = "Embossed" & CStr(I) & ".ipt"
        Dim strFullFileName As String
        strFullFileName = strFilePath & "\" & fileName
        Return strFullFileName
    End Function

    Function adjust(fileName As String, theta As Double) As Double
        openFile(createFileName(fileName))
        calculateGain(theta, "angulo")
        sp = theta
        Try
            changeParameter(openMainSketch(oPartDoc))
            checkBuilder("angulo")

        Catch ex As Exception
            Call UndoCommand()
            makeallDriven()
        End Try

        Return oTheta._Value
    End Function
    Function openMainSketch(oDoc As PartDocument) As Sketch3D

        oSk3D = oDoc.ComponentDefinition.Sketches3D.Item("MainSketch")
        Return oSk3D
    End Function
    Public Sub changeParameter(oSk3D As Sketch3D)
        Try

            oTheta = getParameter("angulo")

            oSk3D.Edit()
            p = iterate("angulo", sp)
            oSk3D.Solve()
            oSk3D.ExitEdit()

        Catch ex4 As Exception
            Call UndoCommand()
            makeallDriven()
            resolution = 1000
            MsgBox(ex4.ToString())
            MsgBox("Fail Iteration  last value: " & oTheta.Value.ToString)
            Exit Sub
        End Try


    End Sub

    Public Sub makeallDriven()
        getDimension("techo").Driven = True
        getDimension("doblez").Driven = True
        getDimension("foldez").Driven = True
        getDimension("gap2").Driven = True
        getDimension("gap1").Driven = True
        'getDimension("angulo").Driven = True

        'getDimension("techo").Driven = True

    End Sub
    Function getParameter(name As String) As Parameter
        Try
            p = oPartDoc.ComponentDefinition.Parameters.ModelParameters.Item(name)
        Catch ex As Exception
            Try
                p = oPartDoc.ComponentDefinition.Parameters.ReferenceParameters.Item(name)
            Catch ex1 As Exception
                Try
                    p = oPartDoc.ComponentDefinition.Parameters.UserParameters.Item(name)
                Catch ex2 As Exception
                    MsgBox(ex2.ToString())
                    MsgBox("Parameter not found: " & name)
                End Try

            End Try

        End Try

        Return p
    End Function

    Function calculateGain(setValue As Double, name As String) As Double
        p = getParameter(name)
        delta = (setValue - p._Value) / (setValue * resolution)
        gain = Math.Exp(delta)
        Return gain
    End Function
    Public Sub UndoCommand()
        Dim oCommandMgr As CommandManager
        oCommandMgr = oApp.CommandManager

        ' Get control definition for the line command. 
        Dim oControlDef As ControlDefinition
        oControlDef = oCommandMgr.ControlDefinitions.Item("AppUndoCmd")
        ' Execute the command. 
        Call oControlDef.Execute()
    End Sub
    Public Sub checkGaps(name As String)

        Try
            If Not name = "techo" Then
                checkAngulos("techo", 2.8, 3.1, Math.Max(2.8, getParameter("techo")._Value * 0.9))
            End If

            If Not name = "gap1" Then
                checkDimension("gap1", 3, 12, 10 * Math.Max(0.3, getParameter("gap1")._Value / 2))
            End If
            If Not name = "gap2" Then
                checkDimension("gap2", 2, 12, 10 * Math.Max(0.2, getParameter("gap2")._Value / 2))
            End If

            If Not name = "foldez" Then
                checkDimension("foldez", getParameter("doblez")._Value * 10, 3, 10 * Math.Max(getParameter("doblez")._Value, getParameter("foldez")._Value / 2))
            End If

            If Not name = "doblez" Then
                checkDimension("doblez", 0.1, 1.2, 10 * Math.Max(0.01, getParameter("doblez")._Value / 2))
            End If



        Catch ex As Exception
            UndoCommand()
            makeallDriven()
            resolution = 1000
            MsgBox(ex.ToString())
            MsgBox("Fail Iteration  checkGaps: " & p.Value.ToString)
        End Try

    End Sub
    Public Function iterate(name As String, setpoint As Double) As Parameter
        p = getParameter(name)

        Try
            calculateGain(setpoint, name)
            makeallDriven()
            getDimension(name).Driven = False
            While (Math.Abs(delta * resolution)) > (setpoint * 10 / resolution)
                p = getParameter(name)
                p.Value = p.Value * calculateGain(setpoint, name)
                checkBuilder(p.Name)

            End While
            getDimension("techo").Driven = False
        Catch ex As Exception
            UndoCommand()
            getDimension(name).Driven = True
            resolution = 1000
            MsgBox(ex.ToString())
            MsgBox("Fail Iteration  last value: " & p.Value.ToString)
            Return p

        End Try
        'getDimension(name).Driven = True
        resolution = 100

        Return p
    End Function

    Public Sub checkBuilder(name As String)
        Try
            oSk3D.Solve()
            oSk3D.ExitEdit()
            oPartDoc.Update()
            oSk3D.Edit()
            checkGaps(name)

        Catch ex As Exception
            UndoCommand()
            makeallDriven()
            getDimension(p.Name).Driven = False
            resolution = 1000
            MsgBox(ex.ToString())
            MsgBox("Fail checkbuilder last value: " & p.Value.ToString)
            Exit Sub
        End Try

    End Sub
    Public Function checkDimension(name As String, a As Double, b As Double, setpoint As Double) As Parameter
        p = getParameter(name)

        If (p._Value < a / 10 Or p._Value > b / 10) Then
            getDimension(name).Driven = False
            p = iterate(name, setpoint / 10)
            getDimension(name).Driven = True
        End If

        Return p
    End Function
    Public Function checkAngulos(name As String, a As Double, b As Double, setpoint As Double) As Parameter
        p = getParameter(name)

        If (p._Value < a Or p._Value > b) Then
            getDimension(name).Driven = False
            p = iterate(name, setpoint)
            getDimension(name).Driven = True
        End If

        Return p
    End Function

    Function getDimension(name As String) As DimensionConstraint3D
        For Each dimension In oSk3D.DimensionConstraints3D
            If dimension.Parameter.Name = name Then
                Return dimension
            End If
        Next
        Return Nothing
    End Function



End Class
