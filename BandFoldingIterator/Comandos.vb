Imports Inventor
Public Class Comandos
    Dim oCommandMgr As CommandManager
    Dim oControlDef As ControlDefinition

    Public Sub New(App As Inventor.Application)
        oCommandMgr = App.CommandManager
    End Sub

    Public Sub UndoCommand()


        ' Get control definition for the line command. 

        oControlDef = oCommandMgr.ControlDefinitions.Item("AppUndoCmd")

        ' Execute the command. 
        Call oControlDef.Execute()
    End Sub
    Function IsUndoable() As Boolean
        Dim ud As Boolean
        oControlDef = oCommandMgr.ControlDefinitions.Item("AppUndoCmd")
        ud = oControlDef.Enabled
        Return ud
    End Function
End Class
