'This iLogic Code Centre's and Arranges Dimensions
'Code by @ClintBrown3D 'Originally posted at https://clintbrown.co.uk/ilogic-centre-arrange-dimensions

On Error GoTo ClintBrown3D
'Code to Centre Dimensions - Adapted from the Inevntor API Sample by @ClintBrown3D
	' a reference to the active drawing document
    Dim oDoc As DrawingDocument
    oDoc = ThisApplication.ActiveDocument

    ' a reference to the active sheet & dimension
    Dim oSheet As Sheet 
    oSheet = oDoc.ActiveSheet
    Dim oDrawingDim As DrawingDimension

    ' Iterate over all dimensions in the drawing and center them if they are linear or angular.
    For Each oDrawingDim In oSheet.DrawingDimensions
        If TypeOf oDrawingDim Is LinearGeneralDimension Or TypeOf oDrawingDim Is AngularGeneralDimension Then
            Call oDrawingDim.CenterText
        End If
    Next
'--------------------------------------------------------------------------------------------------------------------	
'Code to Centre Dimensions - Adapted from https://modthemachine.typepad.com/my_weblog/2009/03/running-commands-using-the-api.html
' Get the active document, assuming it is a drawing.
    Dim oDrawDoc As DrawingDocument
    oDrawDoc = ThisApplication.ActiveDocument

    ' Get the collection of dimensions on the active sheet.
    Dim oDimensions As DrawingDimensions
    oDimensions = oDrawDoc.ActiveSheet.DrawingDimensions
	
	    ' Get a reference to the select set and clear it.
    Dim oSelectSet As SelectSet
    oSelectSet = oDrawDoc.SelectSet
    oSelectSet.Clear

    ' Add each dimension to the select set to select them.
    Dim oDrawDim As DrawingDimension
    For Each oDrawDim In oDimensions
        oSelectSet.Select(oDrawDim)
    Next
    	
	Call ThisApplication.CommandManager.ControlDefinitions.Item("DrawingArrangeDimensionsCmd").Execute
	
Return
ClintBrown3D :
MessageBox.Show("We've encountered a mystery", "Unofficial Inventor", MessageBoxButtons.OK, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1)