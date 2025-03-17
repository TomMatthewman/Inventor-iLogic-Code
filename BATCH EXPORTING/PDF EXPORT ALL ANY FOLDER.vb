Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Timers

Sub Main()
    Dim oDoc As Document = ThisDoc.Document

    ' Ensure the script is run from an assembly or drawing document
    If Not (oDoc.DocumentType = kAssemblyDocumentObject Or oDoc.DocumentType = kDrawingDocumentObject) Then
        MessageBox.Show("Please run this rule from the assembly or drawing files.", "iLogic")
        Exit Sub
    End If

    ' Get user confirmation
    If MessageBox.Show( _
        "This will export all of the files referenced by this document that have drawings as PDFs." _
        & vbLf & "This rule expects that the drawing file shares the same name and is stored in drive D" _
        & vbLf & " " _
        & vbLf & "Are you sure you want to export Drawings for all of the referenced documents?" _
        & vbLf & "This could take a while.", "iLogic - Batch Output PDFs", MessageBoxButtons.YesNo) = vbNo Then
        Exit Sub
    End If

    ' Initialize collections
    Dim missingDrawings As New HashSet(Of String)()
    Dim exportedDrawings As New HashSet(Of String)()
    Dim directoryPath As String = "M:\From Drive D" ' Directory for searching drawing files
    Dim assemblyName As String = System.IO.Path.GetFileNameWithoutExtension(oDoc.FullFileName) ' Get the assembly name
    Dim outputDirectoryPath As String = System.IO.Path.Combine("M:\Exported PDFs", assemblyName & " PDF Export") ' Create a specific folder for this assembly

    ' Ensure output directory exists
    If Not System.IO.Directory.Exists(outputDirectoryPath) Then
        System.IO.Directory.CreateDirectory(outputDirectoryPath)
    End If

    ' Scan for drawing files asynchronously with a timeout
    ScanDrawingFilesInBackground(directoryPath, outputDirectoryPath, missingDrawings, exportedDrawings, oDoc)
End Sub

' Scan and map drawing files asynchronously using Tasks
Sub ScanDrawingFilesInBackground(directoryPath As String, outputDirectoryPath As String, missingDrawings As HashSet(Of String), exportedDrawings As HashSet(Of String), oDoc As Document)
    Dim drawingFiles As New Dictionary(Of String, String)
    Dim cancellationTokenSource As New CancellationTokenSource()

    ' Define the Task to scan files
    Dim task As Task = Task.Run(Sub()
                                    ' Scan drawing files
                                    Dim files As String() = System.IO.Directory.GetFiles(directoryPath, "*.idw", SearchOption.AllDirectories)
                                    For Each file As String In files
                                        Dim drawingFileName As String = System.IO.Path.GetFileNameWithoutExtension(File)
                                        If Not drawingFiles.ContainsKey(drawingFileName) Then
                                            drawingFiles.Add(drawingFileName, File)
                                        End If
                                    Next
                                End Sub, cancellationTokenSource.Token)

    ' Timer for 1-minute timeout
    Dim timeoutTimer As New Timers.Timer(60000) ' 1 minute
    AddHandler timeoutTimer.Elapsed, Sub(sender, e)
                                         If Not task.IsCompleted Then
                                             cancellationTokenSource.Cancel() ' Cancel the task if it times out
                                         End If
                                         timeoutTimer.Stop()
                                     End Sub

    ' Start the timer
    timeoutTimer.AutoReset = False
    timeoutTimer.Start()

    ' When the task is done, export the documents as PDFs
    task.ContinueWith(Sub()
                          If Not task.IsCanceled Then
                              ' Start exporting the documents now
                              ExportDrawingAsPDF(oDoc.FullFileName, drawingFiles, outputDirectoryPath, missingDrawings, exportedDrawings)
                              Dim topLevelDocs As List(Of Document) = GetTopLevelReferencedDocuments(oDoc)
                              For Each doc In topLevelDocs
                                  ExportDocumentAndSubComponents(doc, drawingFiles, outputDirectoryPath, missingDrawings, exportedDrawings)
                              Next
                              DisplayMissingDrawings(missingDrawings)

                              ' Open the output directory after exporting
                              Process.Start("explorer.exe", outputDirectoryPath)
                          Else
                              MsgBox("Drawing file search was cancelled or timed out.")
                          End If
                      End Sub)
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

' Export the document and its sub-components if it's a sub-assembly
Sub ExportDocumentAndSubComponents(ByVal oDoc As Document, ByVal drawingFiles As Dictionary(Of String, String), outputDirectoryPath As String, missingDrawings As HashSet(Of String), exportedDrawings As HashSet(Of String))
    ' Export the document drawing as PDF first
    ExportDrawingAsPDF(oDoc.FullFileName, drawingFiles, outputDirectoryPath, missingDrawings, exportedDrawings)

    ' If the document is an assembly, export its sub-components
    If oDoc.DocumentType = kAssemblyDocumentObject Then
        Dim subDocs As New List(Of Document)
        For Each oRefDoc As Document In oDoc.AllReferencedDocuments
            If oRefDoc.DocumentType = kAssemblyDocumentObject Or oRefDoc.DocumentType = kPartDocumentObject Then
                subDocs.Add(oRefDoc)
            End If
        Next
        subDocs.Sort(Function(x, y) String.Compare(System.IO.Path.GetFileNameWithoutExtension(x.FullFileName), System.IO.Path.GetFileNameWithoutExtension(y.FullFileName)))
        For Each subDoc In subDocs
            ExportDocumentAndSubComponents(subDoc, drawingFiles, outputDirectoryPath, missingDrawings, exportedDrawings)
        Next
    End If
End Sub

' Export a drawing as a PDF if it exists
Sub ExportDrawingAsPDF(ByVal filePath As String, ByVal drawingFiles As Dictionary(Of String, String), outputDirectoryPath As String, missingDrawings As HashSet(Of String), exportedDrawings As HashSet(Of String))
    Dim partFileName As String = System.IO.Path.GetFileNameWithoutExtension(filePath)
    
    ' Skip exporting if already exported
    If exportedDrawings.Contains(partFileName) Then
        Return
    End If

    ' Find the corresponding drawing file
    Dim drawingFilePath As String = FindDrawing(partFileName, drawingFiles)

    ' If drawing file is found, open, export as PDF, and close it
    If Not String.IsNullOrEmpty(drawingFilePath) Then
        Dim oDrawDoc As DrawingDocument = ThisApplication.Documents.Open(drawingFilePath, True)
        InventorVb.DocumentUpdate()
        ExportDrawingDocumentAsPDF(oDrawDoc, outputDirectoryPath, partFileName)
        ' Close the drawing document without saving changes
        oDrawDoc.Close(True)
        exportedDrawings.Add(partFileName) ' Mark as exported
    Else
        missingDrawings.Add(filePath) ' Add to missing drawings list
    End If
End Sub

' Export the drawing document as a PDF
Sub ExportDrawingDocumentAsPDF(ByVal oDrawDoc As DrawingDocument, ByVal outputDirectoryPath As String, ByVal partFileName As String)
    Dim oPDFAddIn As TranslatorAddIn = ThisApplication.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
    Dim oTranslationContext As TranslationContext = ThisApplication.TransientObjects.CreateTranslationContext
    oTranslationContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

    Dim oOptions As NameValueMap = ThisApplication.TransientObjects.CreateNameValueMap
    If oPDFAddIn.HasSaveCopyAsOptions(oDrawDoc, oTranslationContext, oOptions) Then
        oOptions.Value("All_Color_AS_Black") = False
        oOptions.Value("Remove_Line_Weights") = False
    End If

    ' Create the PDF filename and path
    Dim pdfFileName As String = System.IO.Path.GetFileNameWithoutExtension(oDrawDoc.FullFileName) & ".pdf"
    Dim pdfFilePath As String = System.IO.Path.Combine(outputDirectoryPath, pdfFileName)

    ' Create the DataMedium object for the PDF file
    Dim oDataMedium As DataMedium = ThisApplication.TransientObjects.CreateDataMedium
    oDataMedium.FileName = pdfFilePath

    ' Export the document to PDF
    oPDFAddIn.SaveCopyAs(oDrawDoc, oTranslationContext, oOptions, oDataMedium)
End Sub

' Find the drawing file for a given part file name
Function FindDrawing(ByVal partFileName As String, ByVal drawingFiles As Dictionary(Of String, String)) As String
    If drawingFiles.ContainsKey(partFileName) Then
        Return drawingFiles(partFileName)
    Else
        Return ""
    End If
End Function

' Display a message box with the list of missing drawings
Sub DisplayMissingDrawings(ByVal missingDrawings As HashSet(Of String))
    Dim message As String
    If missingDrawings.Count = 0 Then
        message = "Export complete."
    Else
        message = "Export complete." & vbLf & "Missing drawings:" & vbLf & String.Join(vbLf, missingDrawings)
    End If
    MsgBox(message)
End Sub
