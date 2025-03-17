Sub Main()
  Dim myDate As String = Now().ToString("yyyy-MM-dd HHmmss")
  myDate = myDate.Replace(":","")  ' & " - " & TypeString
 userChoice = InputRadioBox("Defined the scope", "This Document", "All Open Documents", True, Title := "Defined the scope")
 UserSelectedActionList = New String(){"DXF & PDF", "PDF Only", "DXF Only"}
  UserSelectedAction = InputListBox("What action must be performed with selected views?", _
          UserSelectedActionList, UserSelectedActionList(0), Title := "Action to Perform", ListName := "Options")
      Select UserSelectedAction
   Case "DXF & PDF": UserSelectedAction = 3
   Case "PDF Only": UserSelectedAction = 1
   Case "DXF Only":    UserSelectedAction = 2
   End Select
 If userChoice Then
   Call MakePDFFromDoc(ThisApplication.ActiveDocument, myDate, UserSelectedAction)
  Else
   For Each oDoc In ThisApplication.Documents
    If oDoc.DocumentType = kDrawingDocumentObject
     Try
      If Len(oDoc.File.FullFileName)>0 Then
       Call MakePDFFromDoc(oDoc, myDate, UserSelectedAction)
      End If
     Catch
     End Try
    End If
   Next
  End If
 End Sub
 Sub MakePDFFromDoc(ByRef oDocument As Document, DateString As String, UserSelectedAction As Integer)
 ' oPath = oDocument.Path
 ' oFileName = oDocument.FileName(False) 'without extension
  'oDocument = ThisApplication.ActiveDocument
  oPDFAddIn = ThisApplication.ApplicationAddIns.ItemById _
  ("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
  oContext = ThisApplication.TransientObjects.CreateTranslationContext
  oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
  oOptions = ThisApplication.TransientObjects.CreateNameValueMap
  oDataMedium = ThisApplication.TransientObjects.CreateDataMedium
 oFullFileName = oDocument.File.FullFileName
  oPath = Left(oFullFileName, InStrRev(oFullFileName, "\")-1)
  oFileName = Right(oFullFileName, Len(oFullFileName)-InStrRev(oFullFileName, "\"))
  oFilePart = Left(oFileName, InStrRev(oFileName, ".")-1)
 'oRevNum = oDocument.iProperties.Value("Project", "Revision Number")
  'oDocument = ThisApplication.ActiveDocument
 ' If oPDFAddIn.HasSaveCopyAsOptions(oDataMedium, oContext, oOptions) Then
  oOptions.Value("All_Color_AS_Black") = 0
  oOptions.Value("Remove_Line_Weights") = 0
  oOptions.Value("Vector_Resolution") = 400
  oOptions.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets
  'oOptions.Value("Custom_Begin_Sheet") = 2
  'oOptions.Value("Custom_End_Sheet") = 4
 ' End If
 'get PDF target folder path
  'oFolder = Left(oPath, InStrRev(oPath, "\")) & "PDF"
  oFolder = oPath & "\iLogic PDF's (" & DateString & ")"
 'Check for the PDF folder and create it if it does not exist
  If Not System.IO.Directory.Exists(oFolder) Then
   System.IO.Directory.CreateDirectory(oFolder)
  End If
 'Set the PDF target file name
  oDataMedium.FileName = oFolder & "\" & oFilePart & ".pdf"
 'Publish document
  If (UserSelectedAction = 1) Or (UserSelectedAction = 3) Then
   oPDFAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)'For PDF's
  End If
  If (UserSelectedAction = 2) Or (UserSelectedAction = 3) Then
   oDocument.SaveAs(oFolder & "\" & oFilePart & ".dxf", True) 'For DXF's
  End If
  'oDocument.SaveAs(oFolder & "\" & ThisDoc.ChangeExtension(".dxf"), True) 'For DXF's
  '------end of iLogic-------
 End Sub