Attribute VB_Name = "_AddInAPI"
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Function: StartAddIn
'---------------------------------------------------------------------------------------
'
' Default start of add-in ... open add-in form
'
'---------------------------------------------------------------------------------------
Public Function StartAddIn()
   StartApplication
End Function


'---------------------------------------------------------------------------------------
' Function: RunVcsCheckDialog
'---------------------------------------------------------------------------------------
'
'  Equal to RunVcsCheck(True)
'
'---------------------------------------------------------------------------------------
Public Function RunVcsCheckDialog() As Variant
   RunVcsCheckDialog = RunVcsCheck(True)
End Function


'---------------------------------------------------------------------------------------
' Function: RunVcsCheck
'---------------------------------------------------------------------------------------
'
' Compare lettercase from CurrentVbProject with saved (table/file) dictionary items
'
' Parameters:
'     OpenDialogToFixLettercase - (Boolean) - Open dialog to fix lettercase
'
' Returns:
'      Boolean (True) ... if DiffCount = 0
'      String         ... if DiffCount > 0 => "Failed: <lettercase info>"
'
'---------------------------------------------------------------------------------------
Public Function RunVcsCheck(Optional ByVal OpenDialogToFixLettercase As Boolean = False) As Variant

   Dim DictFilePath As String
   Dim CheckMsg As String
   Dim DiffCnt As Long
   Dim UseTable As Boolean
   Dim StoreDictData As Boolean
   Dim IntialCnt As Long

   With New DeclarationDict

      If .LoadFromTable(DefaultDeclDictTableName) Then
         UseTable = True
      Else
         DictFilePath = CurrentProject.Path & "\" & CurrentProject.Name & ".DeclarationDict.txt"
         If Not .LoadFromFile(DictFilePath) Then
            .ImportVBProject CurrentVbProject
            ' ... log info: first export
            .ExportToFile DictFilePath
            RunVcsCheck = "Info: no export data exists, run first export"
            Exit Function
         End If
      End If

      IntialCnt = .Count
      .ImportVBProject CurrentVbProject

      DiffCnt = .DiffCount

      If DiffCnt = 0 Then
         If .Count <> IntialCnt Then
            StoreDictData = True
         End If
      End If

      If OpenDialogToFixLettercase Then
         If DiffCnt > 0 Then
            SetDeclarationDictTransferReference .Self
            DoCmd.OpenForm "DeclarationDictApiDialog", , , , , acDialog
            DiffCnt = .DiffCount
            If DiffCnt = 0 Then
               StoreDictData = True
            End If
         End If
      End If

      If StoreDictData Then
         If UseTable Then
            .SaveToTable DefaultDeclDictTableName
         Else
            .ExportToFile DictFilePath
         End If
      End If

      If DiffCnt > 0 Then
         CheckMsg = .DiffCount & " word" & IIf(.DiffCount > 1, "s", vbNullString) & " with different letter case"
         RunVcsCheck = "Failed: " & CheckMsg
      Else
         RunVcsCheck = True
      End If

   End With

End Function
