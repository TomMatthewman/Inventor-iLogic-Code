'iLogic code to calculate mass of selected components
Dim oAsm As AssemblyDocument = ThisDoc.Document
Dim oSelSet As SelectSet = oAsm.SelectSet
Dim UoM As UnitsOfMeasure = oAsm.UnitsOfMeasure
Dim totMass As Double = 0

For Each oObj As Object In oSelSet
    If TypeOf (oObj) Is ComponentOccurrence Then
        Dim oOcc As ComponentOccurrence = oObj
        totMass = totMass + oOcc.MassProperties.Mass
    End If
Next

MessageBox.Show("Total mass of selected occurrences: " & _
UoM.GetStringFromValue(totMass, UoM.MassUnits), "Total mass")



