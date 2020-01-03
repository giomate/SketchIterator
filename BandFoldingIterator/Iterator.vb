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
    Dim fileName, pattern As String
    Dim folding As BandRefolding
    Dim newName As Nombres

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
            done = StartIterator()
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
                fileName = "BandFoldingIteration0.ipt"
                pattern = "Iteration"
                newname = New Nombres
                folding = New BandRefolding(Banda.CreateFileCopy(fileName, newName.IncrementLabelIpt(fileName, pattern)))
                b = folding.StartFolding(0.77)

            End If

        Catch ex As Exception
            Debug.Print(ex.ToString())
            b = False
        End Try
        Return b
    End Function

End Class
