'Get the active document and check whether it's drawing document
If ThisApplication.ActiveDocument.DocumentType = kDrawingDocumentObject Then
Dim oDrgDoc As DrawingDocument
oDrgDoc = ThisApplication.ActiveDocument

' Set reference to drawing print manager
Dim oDrgPrintMgr As DrawingPrintManager
oDrgPrintMgr = oDrgDoc.PrintManager

' Set the printer name
' comment line below to use default printer
oDrgPrintMgr.Printer = "\\server03\Drawing Printer"
'Set the Scale and orientation 
oDrgPrintMgr.ScaleMode = kPrintBestFitScale
'oDrgPrintMgr.ScaleMode = kPrintCurrentWindow

'Set the paper size
oDrgPrintMgr.PaperSize = kPaperSizeA4
'kPaperSizeA3
'oDrgPrintMgr.PaperSize = kPaperSizeA2
'oDrgPrintMgr.PaperSize = kPaperSizeA1
'oDrgPrintMgr.PaperSize = kPaperSizeA0

'Set the Range - Current or All - To use this Comment out " Sheet Range" section below
oDrgPrintMgr.PrintRange = kPrintAllSheets
'oDrgPrintMgr.PrintRange = kPrintCurrentSheet

'Set Sheet Range - To use this Comment out "Range" section above
'Dim FromSheet As Integer = 1
'Dim ToSheet As Integer = 2
'oDrgPrintMgr.PrintRange = kPrintSheetRange
'oDrgPrintMgr.SetSheetRange(FromSheet, ToSheet)

'Set Colours - True or False
oDrgPrintMgr.AllColorsAsBlack = False
'oDrgPrintMgr.AllColorsAsBlack = True

'Set Lineweights Behaviour
oDrgPrintMgr.RemoveLineWeights = False
'oDrgPrintMgr.RemoveLineWeights = True

'Set Rotation Behaviour
oDrgPrintMgr.Rotate90Degrees = False
'oDrgPrintMgr.Rotate90Degrees = True
'Set Tiling behaviour
oDrgPrintMgr.TilingEnabled = False
'oDrgPrintMgr.TilingEnabled = True

'Set the Orientation
oDrgPrintMgr.Orientation = kLandscapeOrientation
'oDrgPrintMgr.Orientation = kPortraitOrientation
oDrgPrintMgr.SubmitPrint
End If
