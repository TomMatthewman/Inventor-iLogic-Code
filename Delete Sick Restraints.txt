Dim oAssDoc As AssemblyDocument
oAssDoc = ThisApplication.ActiveDocument
Dim oConstraint As AssemblyConstraint

RUSure = MessageBox.Show _
("Are you sure you want to Delete all sick constraints?",  _
"iLogic",MessageBoxButtons.YesNo)

If RUSure = vbNo Then
Return
Else
          i = 0
          For Each oConstraint In oAssDoc.ComponentDefinition.Constraints
            If oConstraint.HealthStatus <> oConstraint.HealthStatus.kUpToDateHealth And _
            oConstraint.HealthStatus <> oConstraint.HealthStatus.kSuppressedHealth Then
          oConstraint.Delete
            i = i + 1
          End If
          Next
End If
MessageBox.Show(" A total of "&  i & " constraints were deleted.", "iLogic")