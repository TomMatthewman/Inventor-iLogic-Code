Sub Main()
    Dim oDoc As Document = ThisDoc.Document

    ' Ensure the script is run from an assembly or drawing document
    If Not (oDoc.DocumentType = kAssemblyDocumentObject Or oDoc.DocumentType = kDrawingDocumentObject) Then
        MessageBox.Show("Please run this rule from the assembly or drawing files.", "iLogic")
        Exit Sub
    End If

    ' Get user confirmation
    If MessageBox.Show( _
        "This will Print all of the files referenced by this document that have drawings files." _
        & vbLf & "This rule expects that the drawing file shares the same name and is stored in drive D" _
        & vbLf & " " _
        & vbLf & "Are you sure you want to Print Drawings for all of the referenced documents?" _
        & vbLf & "This could take a while.", "iLogic - Batch Output PDFs", MessageBoxButtons.YesNo) = vbNo Then
        Exit Sub
    End If

    ' Initialize collections
    Dim missingDrawings As New HashSet(Of String)()
    Dim printedDrawings As New HashSet(Of String)()
    Dim directoryPath As String = "M:\From Drive D"

    ' Scan for drawing files in the specified directory
    Dim drawingFiles As Dictionary(Of String, String) = ScanDrawingFiles(directoryPath)

    ' Print the main assembly drawing first
    PrintDrawing(oDoc.FullFileName, drawingFiles, missingDrawings, printedDrawings)

    ' Get all top-level referenced documents and sort them
    Dim topLevelDocs As List(Of Document) = GetTopLevelReferencedDocuments(oDoc)

    ' Print the top-level documents in sorted order
    For Each doc In topLevelDocs
        PrintDocumentAndSubComponents(doc, drawingFiles, missingDrawings, printedDrawings)
    Next

    ' Display the list of missing drawings
    DisplayMissingDrawings(missingDrawings)
End Sub

' Get all top-level referenced documents and sort them alphabetically
Function GetTopLevelReferencedDocuments(ByVal oDoc As Document) As List(Of Document)
    Dim topLevelDocs As New List(Of Document)
    For Each oRefDoc As Document In oDoc.ReferencedDocuments
        If oRefDoc.DocumentType = kAssemblyDocumentObject Or oRefDoc.DocumentType = kPartDocumentObject Then
            topLevelDocs.Add(oRefDoc)
        End If
    Next
    topLevelDocs.Sort(Function(x, y) String.Compare(System.IO.Path.GetFileNameWithoutExtension(x.FullFileName), System.IO.Path.GetFileNameWithoutExtension(y.FullFileName)))
    Return topLevelDocs
End Function

' Print the document and its sub-components if it's a sub-assembly
Sub PrintDocumentAndSubComponents(ByVal oDoc As Document, ByVal drawingFiles As Dictionary(Of String, String), ByRef missingDrawings As HashSet(Of String), ByRef printedDrawings As HashSet(Of String))
    ' Print the document drawing first
    PrintDrawing(oDoc.FullFileName, drawingFiles, missingDrawings, printedDrawings)

    ' If the document is an assembly, print its sub-components
    If oDoc.DocumentType = kAssemblyDocumentObject Then
        Dim subDocs As New List(Of Document)
        For Each oRefDoc As Document In oDoc.AllReferencedDocuments
            If oRefDoc.DocumentType = kAssemblyDocumentObject Or oRefDoc.DocumentType = kPartDocumentObject Then
                subDocs.Add(oRefDoc)
            End If
        Next
        subDocs.Sort(Function(x, y) String.Compare(System.IO.Path.GetFileNameWithoutExtension(x.FullFileName), System.IO.Path.GetFileNameWithoutExtension(y.FullFileName)))
        For Each subDoc In subDocs
            PrintDocumentAndSubComponents(subDoc, drawingFiles, missingDrawings, printedDrawings)
        Next
    End If
End Sub

' Print a drawing if it exists
Sub PrintDrawing(ByVal filePath As String, ByVal drawingFiles As Dictionary(Of String, String), ByRef missingDrawings As HashSet(Of String), ByRef printedDrawings As HashSet(Of String))
    Dim partFileName As String = System.IO.Path.GetFileNameWithoutExtension(filePath)
    
    ' Skip printing if already printed
    If printedDrawings.Contains(partFileName) Then
        Return
    End If

    ' Find the corresponding drawing file
    Dim drawingFilePath As String = FindDrawing(partFileName, drawingFiles)

    ' If drawing file is found, open, print, and close it
    If Not String.IsNullOrEmpty(drawingFilePath) Then
        Dim oDrawDoc As DrawingDocument = ThisApplication.Documents.Open(drawingFilePath, True)
        InventorVb.DocumentUpdate()
        PrintDrawingDocument(oDrawDoc)
        ' Close the drawing document without saving changes
        oDrawDoc.Close(True)
        printedDrawings.Add(partFileName) ' Mark as printed
    Else
        missingDrawings.Add(filePath) ' Add to missing drawings list
    End If
End Sub

' Print the drawing document
Sub PrintDrawingDocument(ByVal oDrawDoc As DrawingDocument)
    Dim oDrgPrintMgr As DrawingPrintManager = oDrawDoc.PrintManager
    With oDrgPrintMgr
        .Printer = "\\server03\Drawing Printer"
        .ScaleMode = kPrintBestFitScale
        .PaperSize = kPaperSizeA4
        .PrintRange = kPrintAllSheets
        .AllColorsAsBlack = False
        .RemoveLineWeights = False
        .Rotate90Degrees = False
        .TilingEnabled = False
        .Orientation = kLandscapeOrientation
        .SubmitPrint()
    End With
End Sub

' Find the drawing file for a given part file name
Function FindDrawing(ByVal partFileName As String, ByVal drawingFiles As Dictionary(Of String, String)) As String
    If drawingFiles.ContainsKey(partFileName) Then
        Return drawingFiles(partFileName)
    Else
        Return ""
    End If
End Function

' Scan and map all drawing files in the specified directory
Function ScanDrawingFiles(ByVal directoryPath As String) As Dictionary(Of String, String)
    Dim drawingFiles As New Dictionary(Of String, String)
    Dim files As String() = System.IO.Directory.GetFiles(directoryPath, "*.idw", System.IO.SearchOption.AllDirectories)

    For Each file As String In files
        Dim drawingFileName As String = System.IO.Path.GetFileNameWithoutExtension(File)
        If Not drawingFiles.ContainsKey(drawingFileName) Then
            drawingFiles.Add(drawingFileName, File)
        End If
    Next

    Return drawingFiles
End Function

' Display a message box with the list of missing drawings
Sub DisplayMissingDrawings(ByVal missingDrawings As HashSet(Of String))
    Dim message As String
    If missingDrawings.Count = 0 Then
        message = "Printing complete."
    Else
        message = "Printing complete." & vbLf & "Missing drawings:" & vbLf & String.Join(vbLf, missingDrawings)
    End If
    MsgBox(message)
End Sub
