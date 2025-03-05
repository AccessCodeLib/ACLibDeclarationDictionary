Attribute VB_Name = "_AddInAPI"
Option Compare Database
Option Explicit

Public Function StartAddIn()

   StartApplication

End Function

Public Function RunVcsCheck() As Variant

   Dim DictFilePath As String

   With New DeclarationDict

      If Not .LoadFromTable(DefaultDeclDictTableName) Then
         DictFilePath = CurrentProject.Path & "\DeclarationDict.txt"
         If Not .LoadFromFile(DictFilePath) Then
            .ImportVBProject CurrentVbProject
            ' ... log info: first export
            Debug.Print "RunVcsCheck: no export data exists, run first export"
            .ExportToFile DictFilePath
            RunVcsCheck = "Info: no export data exists, run first export"
            Exit Function
         End If
      End If

      .ImportVBProject CurrentVbProject

      If .DiffCount > 0 Then
         RunVcsCheck = "Failed: " & .DiffCount & " words with different letter case"
         Debug.Print "RunVcsCheck: " & .DiffCount & " words with different letter case"
      Else
         RunVcsCheck = True
      End If

   End With


End Function
