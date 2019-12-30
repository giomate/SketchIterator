Imports Inventor

Module Conversions

    Public Sub SetUnitsToMetric(Document As Inventor.Document)
        'Get Units of Measure
        Dim oUOM As Inventor.UnitsOfMeasure
        oUOM = Document.UnitsOfMeasure
        'Set length units to Metric and save
        oUOM.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits
        oUOM.MassUnits = UnitsTypeEnum.kGramMassUnits
        oUOM.AngleUnits = UnitsTypeEnum.kRadianAngleUnits
        Document.Dirty = True
        Document.Update()
        'Document.Save()


    End Sub
End Module