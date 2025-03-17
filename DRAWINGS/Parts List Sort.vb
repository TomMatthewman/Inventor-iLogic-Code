On Error Resume Next
Dim oDrawDoc As DrawingDocument
    oDrawDoc = ThisApplication.ActiveDocument
Dim Sheet As Inventor.Sheet
Dim Cursheet As String
     Cursheet = oDrawDoc.ActiveSheet.Name
    
For Each oSheet In oDrawDoc.Sheets
        oSheet.Activate
Dim oPartsList As PartsList
    oPartsList = oDrawDoc.ActiveSheet.PartsLists.Item(1)
        If Not oPartsList Is Nothing Then
         Resume Next
    'Call oPartsList.Sort("ITEM")
	Call oPartsList.Sort("PART NUMBER", 1, "DESCRIPTION", 1, "QTY", 1)
	oPartsList.Renumber
iLogicVb.UpdateWhenDone = True
ActiveSheet = ThisDrawing.Sheet("Sheet:1")
ThisApplication.ActiveView.Fit


End If
    Next	