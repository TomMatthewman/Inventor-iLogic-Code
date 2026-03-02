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


Sub PrintDocumentAndSubComponents(ByVal oDoc As Document, ByVal drawingFiles As Dictionary(Of String, String), ByRef missingDrawings As HashSet(Of String), ByRef printedDrawings As HashSet(Of String))

    PrintDrawing(oDoc.FullFileName, drawingFiles, missingDrawings, printedDrawings)

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


Sub PrintDrawing(ByVal filePath As String, ByVal drawingFiles As Dictionary(Of String, String), ByRef missingDrawings As HashSet(Of String), ByRef printedDrawings As HashSet(Of String))

    Dim partFileName As String = System.IO.Path.GetFileNameWithoutExtension(filePath)

    If printedDrawings.Contains(partFileName) Then Return

    Dim drawingFilePath As String = FindDrawing(partFileName, drawingFiles)

    If Not String.IsNullOrEmpty(drawingFilePath) Then
        Try
            Dim oDrawDoc As DrawingDocument = ThisApplication.Documents.Open(drawingFilePath, True)
            oDrawDoc.Activate()
            ThisApplication.UserInterfaceManager.DoEvents()

            oDrawDoc.Update()
            InventorVb.DocumentUpdate()

            ' Add date annotation
            Dim oSheet As Sheet = oDrawDoc.Sheets(1)
            Dim oGenNotes As DrawingNotes = oSheet.DrawingNotes
            Dim oTG As TransientGeometry = ThisApplication.TransientGeometry

            Dim currentDate As String = DateTime.Now.ToString("dd MMMM yyyy")
            Dim offsetX As Double = 1.0
            Dim offsetY As Double = oSheet.Height - 10.0

            Dim topLeftPoint As Point2d = oTG.CreatePoint2d(offsetX, offsetY)
            Dim oNote As GeneralNote = oGenNotes.GeneralNotes.AddFitted(topLeftPoint, currentDate)

            Dim oColor As Color = ThisApplication.TransientObjects.CreateColor(255, 0, 0)
            oNote.Color = oColor

            oDrawDoc.Update()
            ThisApplication.UserInterfaceManager.DoEvents()

            ' Print
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

            System.Threading.Thread.Sleep(1000)

            oNote.Delete()

            oDrawDoc.Close(True)

            ' Garbage collection cleanup
            GC.Collect()
            GC.WaitForPendingFinalizers()

            printedDrawings.Add(partFileName)

        Catch ex As Exception
            MessageBox.Show("Failed printing: " & drawingFilePath & vbLf & ex.Message)
        End Try
    Else
        missingDrawings.Add(filePath)
    End If
End Sub


Function FindDrawing(ByVal partFileName As String, ByVal drawingFiles As Dictionary(Of String, String)) As String
    If drawingFiles.ContainsKey(partFileName) Then
        Return drawingFiles(partFileName)
    Else
        Return ""
    End If
End Function


Function ScanDrawingFiles(ByVal directoryPath As String) As Dictionary(Of String, String)
    Dim drawingFiles As New Dictionary(Of String, String)
    Dim files As String() = System.IO.Directory.GetFiles(directoryPath, "*.idw", System.IO.SearchOption.AllDirectories)

    For Each file As String In files
        Dim drawingFileName As String = System.IO.Path.GetFileNameWithoutExtension(file)
        If Not drawingFiles.ContainsKey(drawingFileName) Then
            drawingFiles.Add(drawingFileName, file)
        End If
    Next

    Return drawingFiles
End Function


Sub DisplayMissingDrawings(ByVal missingDrawings As HashSet(Of String))
    Dim message As String
    If missingDrawings.Count = 0 Then
        message = "Printing complete."
    Else
        message = "Printing complete." & vbLf & "Missing drawings:" & vbLf & String.Join(vbLf, missingDrawings)
    End If
    MsgBox(message)
End Sub
