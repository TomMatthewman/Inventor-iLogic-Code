'catch and skip errors
On Error Resume Next
'define the active assembly
Dim oAssyDoc As AssemblyDocument
oAssyDoc = ThisApplication.ActiveDocument

'get user input as True or False
wfBoolean = InputRadioBox("Turn all Work Features On/Off", "On", "Off", False, "iLogic")

'Check all referenced docs 
Dim oDoc As Inventor.Document
For Each oDoc In oAssyDoc.AllReferencedDocuments
    'set work plane visibility
    For Each oWorkPlane In oDoc.ComponentDefinition.WorkPlanes
    oWorkPlane.Visible = wfBoolean
    Next
    'set work axis visibility
    For Each oWorkAxis In oDoc.ComponentDefinition.WorkAxes
    oWorkAxis.Visible = wfBoolean
    Next
    'set work point visibility
    For Each oWorkPoint In oDoc.ComponentDefinition.WorkPoints
    oWorkPoint.Visible = wfBoolean
    Next    
    'set sketch visibility
    For Each oSketch In oDoc.ComponentDefinition.Sketches
    oSketch.Visible = wfBoolean
    Next
Next

'update the files
InventorVb.DocumentUpdate()
