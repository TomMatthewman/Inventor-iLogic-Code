' Set a reference to the drawing document.
Dim oDrawDoc As DrawingDocument
    oDrawDoc = ThisApplication.ActiveDocument

'Set a reference to the active sheet.
Dim oSheet As Sheet
    oSheet = oDrawDoc.ActiveSheet
    
Dim oViews as DrawingViews 
	oViews = oSheet.DrawingViews
	
Dim oView As DrawingView

For Each oView In oViews
    Dim oCenterline As AutomatedCenterlineSettings
    'WB added
    oView.GetAutomatedCenterlineSettings(oCenterline)
    'WB moved here
    oCenterline.ApplytoHoles  = True 
    oCenterline.ProjectionParallelAxis = True
    
    oView.SetAutomatedCenterlineSettings(oCenterline)
   ' oCenterline.ApplytoHoles  = True
   ' oCenterline.ProjectionParallelAxis = True
    
'   Dim resultCenters As ObjectsEnumerator
'   resultCenters = oView.SetAutomatedCenterlineSettings(oCenterline)
Next 