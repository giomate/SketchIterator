Imports System.Type
Imports System.Activator
Imports System.Runtime.InteropServices
Imports Inventor


Public Class Iterator
    Dim oApp As Inventor.Application
    Dim oDoc As Inventor.Document
    Dim started, running, done As Boolean
    Dim oDesignProjectMgr As DesignProjectManager
    Dim Banda As InventorFile
    Dim fullFileName, pattern, shortFileName As String
    Dim folding, folded As BandRefolding
    Dim newName As Nombres


    Public Structure ParametersCollection
        Public a As Parameter
        Public b As Parameter
        Public c As Parameter
    End Structure
    Public lados As ParametersCollection
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Try

            oApp = Marshal.GetActiveObject("Inventor.Application")

            started = True

        Catch ex As Exception
            Try
                Dim oInvAppType As Type = GetTypeFromProgID("Inventor.Application")

                oApp = CreateInstance(oInvAppType)
                oApp.Visible = True

                'Note: if you shut down the Inventor session that was started
                'this(way) there is still an Inventor.exe running. We will use
                'this Boolean to test whether or not the Inventor App  will
                'need to be shut down.
                started = True


            Catch ex2 As Exception
                MsgBox(ex2.ToString())
                MsgBox("Unable to get or start Inventor")
            End Try
        End Try

    End Sub

    Private Sub Iterator_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        If started Then
            'oApp.ActiveDocument.Close()
            MsgBox("Closing Session")
        End If
        oApp = Nothing
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            done = StartIterator()
            If done Then
                Me.Close()
            End If
        Catch ex As Exception
            Debug.Print(ex.ToString())
            Debug.Print("Unable to find Document")
        End Try
    End Sub

    Private Sub Iterator_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Try

            done = StartIteratorMatched()
            If done Then
                Me.Close()
            End If
        Catch ex As Exception
            Debug.Print(ex.ToString())
            Debug.Print("Unable to find Document")
        End Try
    End Sub

    Public Function StartIterator() As Boolean
        Dim b As Boolean
        Try
            If (started And (Not running)) Then

                Banda = New InventorFile(oApp)
                fullFileName = "BandFoldingIteration9.ipt"
                pattern = "Iteration"
                newname = New Nombres
                folding = New BandRefolding(Banda.CreateFileCopy(fullFileName, newName.IncrementLabelIpt(fullFileName, pattern)))
                b = folding.StartFolding(0.77)

            End If

        Catch ex As Exception
            Debug.Print(ex.ToString())
            b = False
        End Try
        Return b
    End Function
    Public Function StartIteratorMatched() As Boolean
        Dim b As Boolean

        Try

            If (started And (Not running)) Then

                Banda = New InventorFile(oApp)
                shortFileName = "BandFoldingIteration0.ipt"
                fullFileName = Banda.CreateFullFileName(shortFileName)
                Dim matchFileName As String
                matchFileName = Banda.CreateFullFileName("Band9.ipt")

                pattern = "Iteration"
                newName = New Nombres

                folding = New BandRefolding(Banda.CreateFileCopy(matchFileName, newName.IncrementLabelIpt(fullFileName, pattern)))
                GetInitialParameters(fullFileName)
                folding.lados.a = lados.a
                folding.lados.b = lados.b
                folding.lados.b = lados.c
                'folded.CloseDocument()

                b = folding.StartFoldingAutomatic()

            End If

        Catch ex As Exception
            Debug.Print(ex.ToString())
            b = False
        End Try
        Return b
    End Function
    Function GetInitialParameters(fullname As String) As Integer
        folded = New BandRefolding(Banda.OpenFullNameFile(fullname))
        Try
            lados.a = folded.GetParameter("finala")
            lados.b = folded.GetParameter("finalb")
            lados.c = folded.GetParameter("finalc")

            'folded.CloseDocument()
            Return 0
        Catch ex As Exception
            lados.a = folded.GetParameter("finala")
            lados.b = folded.GetParameter("finalb")
            lados.c = folded.GetParameter("finalc")

            'folded.CloseDocument()
            Return 0
        End Try

    End Function

End Class
