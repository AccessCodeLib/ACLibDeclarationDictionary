Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_DeclarationDict As DeclarationDict

Private Sub cmdLoadFromTable_Click()
   LoadFromTable
End Sub

Private Sub cmdSaveToTable_Click()
   SaveToTable
End Sub

Private Sub Form_Load()

   With CurrentApplication
      Me.lblVersionInfo.Caption = .ApplicationFullName & " " & ChrW(&H2022) & " Version " & CurrentApplication.Version
   End With

   CodeDb.Execute "delete from tabWords"
   Set m_DeclarationDict = New DeclarationDict

End Sub

Private Sub Form_Timer()
   Me.TimerInterval = 0
   Me.lblTableRecInfo.Caption = vbNullString
   Me.lblVersionInfo.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   CodeDb.Execute "delete from tabWords"
End Sub

Private Sub cmdChangeLetterCase_Click()
    ChangeLetterCase
End Sub

Private Sub cmdUpdateDict_Click()
   m_DeclarationDict.ImportVBProject CurrentVbProject
   RefreshList
End Sub

Private Sub lbDictData_AfterUpdate()
   FillWordVariationsList
End Sub

Private Sub lbVariations_AfterUpdate()
   SetVariationCommitMode
End Sub

Private Sub ogViewMode_AfterUpdate()
   RefreshList
End Sub

Private Sub RefreshList()

   CodeDb.Execute "delete from tabWords"
   Me.lbDictData.Requery

   Dim rs As DAO.Recordset
   Set rs = CodeDb.OpenRecordset("tabWords", dbOpenDynaset, dbAppendOnly)

   Dim listDict As Scripting.Dictionary
   Set listDict = m_DeclarationDict.ToDict(CBool(Me.ogViewMode.Value))

   Dim DictIndex As Long
   Dim DictKey As String

   For DictIndex = 0 To listDict.Count - 1

      DictKey = listDict.Keys(DictIndex)

      rs.AddNew
      rs.Fields(0).Value = DictKey
      rs.Fields(1).Value = listDict.Item(DictKey)
      rs.Update

   Next

   RequeryDictData

End Sub

Private Sub RequeryDictData()

   ResetVariations

   Me.txtCurrentLetterCase.Value = Null
   Me.lbDictData.Value = Null
   With Me.lbDictData
      .Value = Null
      If Len(.RowSource) = 0 Then
         .RowSource = "SELECT W.Word, W.Variations FROM tabWords AS W ORDER BY W.Word;"
      Else
         .Requery
      End If
   End With

   RefreshDictInfo

End Sub

Private Sub RefreshDictInfo()
   Me.txtDictInfo.Value = m_DeclarationDict.DiffCount & " / " & m_DeclarationDict.Count
End Sub

Private Sub ResetVariations()
   Me.lbVariations.RowSource = vbNullString
   SetVariationCommitMode
End Sub

Private Sub FillWordVariationsList()

   Dim lb As ListBox
   Set lb = Me.lbVariations
   lb.Value = Null
   lb.RowSource = vbNullString

   Dim Items() As String
   Items = m_DeclarationDict.GetWordVariants(Me.txtWord.Value)

   Dim i As Long
   For i = LBound(Items) To UBound(Items)
      lb.AddItem Items(i)
   Next

   Me.txtCurrentLetterCase.Value = Items(UBound(Items))

   SetVariationCommitMode

End Sub

Private Sub SetVariationCommitMode()

   Dim EnableChange As Boolean

   EnableChange = IsDifferentLetterCase(Nz(Me.txtCurrentLetterCase.Value, vbNullString), Nz(Me.txtSelectedLetterCase.Value, vbNullString))
   If Not EnableChange Then
      EnableChange = IsDifferentLetterCase(Nz(Me.txtWord.Value, vbNullString), Nz(Me.txtSelectedLetterCase.Value, vbNullString))
   End If

   Me.cmdChangeLetterCase.Enabled = EnableChange

End Sub

Private Property Get IsDifferentLetterCase(ByVal String1 As String, ByVal String2 As String) As Boolean

   If Len(String1) = 0 Or Len(String2) = 0 Then
      IsDifferentLetterCase = False
   ElseIf StrComp(String1, String2, vbTextCompare) <> 0 Then
      IsDifferentLetterCase = False
   ElseIf StrComp(String1, String2, vbBinaryCompare) = 0 Then
      IsDifferentLetterCase = False
   Else
      IsDifferentLetterCase = True
   End If

End Property

Private Sub ChangeLetterCase()

    Dim ChangeFrom As String
    Dim ChangeTo As String

    ChangeFrom = Me.txtCurrentLetterCase.Value
    ChangeTo = Me.txtSelectedLetterCase.Value

    If Not IsDifferentLetterCase(ChangeFrom, ChangeTo) Then
      If IsDifferentLetterCase(Me.txtWord.Value, ChangeTo) Then
         SetCurrentLetterCaseAsDefault ChangeTo
      End If
      Exit Sub
    End If

    With New CodeModulGenerator
        .CreateCodemodule vbext_ct_StdModule
        .InsertDeclarationLine "Private " & ChangeTo
        .RemoveCodemodule
    End With

    m_DeclarationDict.AddWord ChangeTo

    FillWordVariationsList

    Me.SetFocus

End Sub

Private Sub SetCurrentLetterCaseAsDefault(ByVal ChangeTo As String)

   m_DeclarationDict.ChangeKeyLetterCase ChangeTo

   Dim UpdateSql As String
   Dim VariantsString As String
   VariantsString = Join(m_DeclarationDict.GetWordVariants(ChangeTo, True), "|")
   With SqlTools
      UpdateSql = "update tabWords set Word = " & .TextToSqlText(ChangeTo) & ", Variations=" & .TextToSqlText(VariantsString) & " where Word = " & .TextToSqlText(ChangeTo)
   End With
   CodeDb.Execute UpdateSql, dbFailOnError

   With Me.lbDictData
      .Requery
      .Value = ChangeTo
   End With

   RefreshDictInfo

End Sub

Private Sub SaveToTable()

   Dim db As DAO.Database
   Set db = CurrentDb

   If Not TableDefExists("USysDeclDict", db) Then
      db.Execute "create table USysDeclDict (DeclWord varchar(255) Not Null Primary Key)", dbFailOnError
   End If

   Dim rs As DAO.Recordset
   Dim i As Long

   db.Execute "delete from USysDeclDict"
   Set rs = db.OpenRecordset("USysDeclDict", dbOpenDynaset, dbAppendOnly)
   For i = 0 To m_DeclarationDict.Count - 1
      rs.AddNew
      rs.Fields(0).Value = m_DeclarationDict.Dict.Keys(i)
      rs.Update
   Next

   ShowTableRecInfo m_DeclarationDict.Count & " records were inserted into USysDeclDict"

End Sub

Private Sub LoadFromTable()

   Dim db As DAO.Database
   Dim rs As DAO.Recordset
   Dim i As Long

   Set db = CurrentDb

   If Not TableDefExists("USysDeclDict", db) Then
      ShowTableRecInfo "Table USysDeclDict not exists"
      Exit Sub
   End If

   Set m_DeclarationDict = New DeclarationDict


   Set rs = db.OpenRecordset("select DeclWord from USysDeclDict", dbOpenSnapshot)
   Do While Not rs.EOF
      m_DeclarationDict.AddWord rs.Fields(0)
      rs.MoveNext
   Loop
   rs.Close
   Set db = Nothing

   ShowTableRecInfo m_DeclarationDict.Count & " records were imported from USysDeclDict"
   RefreshList

End Sub

Private Sub ShowTableRecInfo(ByVal Info As String)
   Me.lblTableRecInfo.Caption = Info
   Me.lblVersionInfo.Visible = False
   Me.TimerInterval = 4000
End Sub
