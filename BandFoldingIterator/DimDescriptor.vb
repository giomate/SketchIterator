Imports Inventor
Public Class DimDescriptor
    Enum DimensionType
        dimension
        angulo
    End Enum
    Public Structure ParametersOptimizer
        Public Name As String
        Public Prio, Setpoint, Ceil, Floor As Double
        Public Type As DimensionType
    End Structure
    Public PO As ParametersOptimizer

    Public Sub New()
        PO.Name = "new"
        PO.Prio = 0
        PO.Setpoint = 0
        PO.Ceil = 0
        PO.Floor = 0
        PO.Type = DimensionType.dimension

    End Sub

    Public Sub setNameParameter(name As String)
        PO.Name = name
    End Sub

    Function GetSetpointDim() As Double
        Dim sp As Double
        sp = PO.Setpoint
        Return sp
    End Function
End Class
