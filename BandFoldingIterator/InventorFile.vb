
Imports Inventor
Public Class InventorFile
    Dim applicacion As Inventor.Application
    Dim documento As Inventor.Document
    Dim started As Boolean
    Public manager As DesignProjectManager
    Public archivador As FileManager



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
        applicacion = App
        manager = applicacion.DesignProjectManager
        documento = applicacion.ActiveDocument
        archivador = applicacion.FileManager

        DP.Dmax = 200
        DP.Dmin = 32
        Tr = (DP.Dmax + DP.Dmin) / 4
        Cr = (DP.Dmax - DP.Dmin) / 4
        DP.p = 17
        DP.q = 37

    End Sub
    Public Function openFile(fileName As String) As PartDocument
        Dim fullname As String = createFileName(fileName)
        Try
            If applicacion.Documents.Count > 0 Then
                If Not (applicacion.ActiveDocument.FullFileName = fullname) Then
                    documento = applicacion.Documents.Open(fullname, True)
                End If
            Else
                documento = applicacion.Documents.Open(fullname, True)
            End If

            documento = applicacion.ActiveDocument
        Catch ex3 As Exception
            Debug.Print(ex3.ToString())
            Debug.Print("Unable to find Document")
        End Try

        ' Conversions.SetUnitsToMetric(oPartDoc)
        Return documento
    End Function

    Public Function OpenFullNameFile(fullName As String) As PartDocument

        Try
            If applicacion.Documents.Count > 0 Then
                If Not (applicacion.ActiveDocument.FullFileName = fullName) Then
                    documento = applicacion.Documents.Open(fullName, True)
                End If
            Else
                documento = applicacion.Documents.Open(fullName, True)
            End If

            documento = applicacion.ActiveDocument
        Catch ex3 As Exception
            Debug.Print(ex3.ToString())
            Debug.Print("Unable to find Document")
        End Try

        ' Conversions.SetUnitsToMetric(oPartDoc)
        Return documento
    End Function
    Public Function CreateFileCopy(fullName As String, saveas As String) As PartDocument

        Dim oPartDoc As PartDocument
        Try
            If applicacion.Documents.Count > 0 Then
                If documento.FullFileName = saveas Then
                    documento.Close(True)
                End If
            End If

            oPartDoc = OpenFullNameFile(fullName)
            oPartDoc.SaveAs(saveas, True)
            oPartDoc.Close(True)
            oPartDoc = OpenFullNameFile(saveas)
            oPartDoc.Update()


        Catch ex As Exception
            Debug.Print(ex.ToString())
            Return Nothing
        End Try
        Return oPartDoc

    End Function

    Public Function CreateFileName(fileName As String) As String
        Dim strFilePath As String
        strFilePath = manager.ActiveDesignProject.WorkspacePath
        ' Dim strFileName As String
        'strFileName = "Embossed" & CStr(I) & ".ipt"
        Dim strFullFileName As String
        strFullFileName = strFilePath & "\" & fileName
        Return strFullFileName
    End Function
    Public Function CreateFullFileName(fileName As String) As String
        Dim strFilePath As String
        strFilePath = manager.ActiveDesignProject.WorkspacePath
        Dim strFullFileName As String
        strFullFileName = strFilePath & "\" & fileName
        Return strFullFileName
    End Function
    Function createFileNameNumber(fileName As String, i As Integer) As String
        Dim strFilePath As String
        strFilePath = manager.ActiveDesignProject.WorkspacePath

        Dim strFullFileName As String
        strFullFileName = strFilePath & "\" & fileName
        Return strFullFileName
    End Function
    Function GetAll() As Integer

        Return 0
    End Function

    Public Function CreateTempAsm() As AssemblyDocument

        Dim oDoc As AssemblyDocument
        oDoc = applicacion.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, applicacion.FileManager.GetTemplateFile(DocumentTypeEnum.kAssemblyDocumentObject), True)

        Return oDoc
    End Function
End Class
