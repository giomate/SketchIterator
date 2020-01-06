Imports Inventor
Public Class DesignDoctor
    Dim app As Inventor.Application
    Dim doc As Inventor.Document
    Dim monitor As DesignMonitoring
    Dim comando As Comandos
    Public Sub New(docu As Inventor.Document)
        doc = docu
        app = docu.Parent
        comando = New Comandos(docu.Parent)
        monitor = New DesignMonitoring(doc)
    End Sub
    Public Function FindBrokenSketch(f As PartFeature) As Sketch3D

        Return Nothing

    End Function
    Public Function UndoFeature(feature As PartFeature) As Boolean
        Try
            While Not monitor.IsFeatureHealthy(feature)
                comando.UndoCommand()
                Debug.Print("!!!! Doctor Undoing !!!")
                doc.Update()
                comando.UndoCommand()
            End While
        Catch ex As Exception
            Debug.Print(ex.ToString())
            Return False
        End Try
        Return True
    End Function

End Class
