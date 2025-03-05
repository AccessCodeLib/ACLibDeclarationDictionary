VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeclarationDict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

#Const DebugPrintEnabled = 0

Private m_Words As Scripting.Dictionary
Private m_WordVariations As Scripting.Dictionary
Private m_DiffCnt As Long

Public Event WordInserted(ByVal WordKey As String)
Public Event WordChanged(ByVal WordKey As String)

Private Sub Class_Initialize()
   Set m_Words = GetNewDict(TextCompare)
   Set m_WordVariations = GetNewDict(TextCompare)
   m_DiffCnt = -1
End Sub

Private Function GetNewDict(ByVal CompareMode As Scripting.CompareMethod) As Scripting.Dictionary
   Dim Dict As Scripting.Dictionary
   Set Dict = New Scripting.Dictionary
   Dict.CompareMode = CompareMode
   Set GetNewDict = Dict
End Function

Friend Property Get VariationsDict() As Scripting.Dictionary
   Set VariationsDict = m_WordVariations
End Property

Public Property Get Count() As Long
   Count = m_Words.Count
End Property

Public Property Get DiffCount() As Long
   If m_DiffCnt = -1 Then
      m_DiffCnt = GetDiffCnt
   End If
   DiffCount = m_DiffCnt
End Property

Private Function GetDiffCnt() As Long

   Dim Word As String
   Dim i As Long
   Dim DifCntr As Long

   For i = 0 To m_Words.Count - 1
      Word = m_Words.Keys(i)
      If IsChangedItem(Word, m_WordVariations.Item(Word)) Then
         DifCntr = DifCntr + 1
      End If
   Next

   GetDiffCnt = DifCntr

End Function

Public Function IsWordWithChangedLetterCase(ByVal Word As String) As Boolean

   IsWordWithChangedLetterCase = IsChangedItem(Word, m_WordVariations.Item(Word))

End Function

Private Function IsChangedItem(ByVal Word As String, ByVal SubDict As Scripting.Dictionary) As Boolean

   Dim LastKey As String

   If StrComp(Word, m_Words.Item(Word), vbBinaryCompare) <> 0 Then
      Word = m_Words.Item(Word)
   End If

   With SubDict
      If .Count > 1 Then
         LastKey = .Keys(.Count - 1)
         If StrComp(LastKey, Word, vbBinaryCompare) <> 0 Then
            IsChangedItem = True
         End If
      End If
   End With

End Function

Public Sub AddWord(ByVal Word As String)

   Word = Trim(Word)

   If Left(Word, 1) = "[" Then
      If Right(Word, 1) = "]" Then
         Word = Trim(Mid(Word, 2, Len(Word) - 2))
      End If
   End If

   If Len(Word) = 0 Then
      Exit Sub
   End If

   m_DiffCnt = -1

   Dim SubDict As Scripting.Dictionary
   Dim LastKey As String
   Dim RaiseInsert As Boolean
   Dim RaiseChanged As Boolean

   If Not m_Words.Exists(Word) Then
      m_Words.Add Word, Word
      m_WordVariations.Add Word, GetNewDict(BinaryCompare)
      RaiseInsert = True
   End If

   Set SubDict = m_WordVariations.Item(Word)
   With SubDict
      If Not .Exists(Word) Then
         .Add Word, Word
         RaiseChanged = True
      ElseIf .Count > 1 Then
         LastKey = .Keys(.Count - 1)
         If StrComp(LastKey, Word, vbBinaryCompare) <> 0 Then
            .Remove Word
            .Add Word, Word
         End If
         RaiseChanged = True
      End If
   End With

   If RaiseInsert Then
      RaiseEvent WordInserted(Word)
   ElseIf RaiseChanged Then
      RaiseEvent WordChanged(Word)
   End If

End Sub

Public Sub ImportVBProject(ByVal VBProject2Import As VBProject)

   Dim vbc As VBComponent

   For Each vbc In VBProject2Import.VBComponents
      ImportVBComponent vbc
   Next

End Sub

Public Sub ImportVBComponent(ByVal VBComponent2Import As VBComponent)
   ImportCodeModule VBComponent2Import.CodeModule
End Sub

Public Sub ImportCodeModule(ByVal CodeModule2Import As CodeModule)
   ImportCode CodeModule2Import.Lines(1, CodeModule2Import.CountOfLines)
End Sub

Public Sub ImportCode(ByVal Code As String)

   Dim RegEx As RegExp
   Set RegEx = NewRegExp

   Code = PrepareCode(Code, RegEx)

   Const ProcIndex As Long = 0
   Const EnumTypeIndex As Long = 1

   Dim Patterns(2) As String

   Patterns(ProcIndex) = "(?:\r|\n|^)\s*(?:Sub|Function|Property Get|Property Let|Property Set|Event)\s+([^\r\n]*)"
   Patterns(EnumTypeIndex) = "(?:\r|\n|^)\s*(?:Enum|Type)([\s\S]*?)(?:End\s+(?:Enum|Type))"
   Patterns(2) = "(?:\r|\n|^)\s*(?:Dim|ReDim|Private|Friend|Public|Const|Static|Global|Implements)\s+([^\r\n]*)"

   Dim i As Long
   For i = 0 To UBound(Patterns)
      RegEx.Pattern = Patterns(i)
      AddFromCode Code, RegEx, i = ProcIndex, i = EnumTypeIndex
   Next

End Sub

Private Function NewRegExp() As RegExp

   Dim RegEx As RegExp

   Set RegEx = New RegExp
   RegEx.IgnoreCase = True
   RegEx.Global = True

   Set NewRegExp = RegEx

End Function

Private Function PrepareCode(ByVal Code As String, ByVal RegEx As RegExp) As String

   Code = Replace(Code, " _" & vbNewLine, " ")

#If DebugPrintEnabled Then
   DebugPrint Code, True, "PrepareCode - after Replace ' _' & vbNewLine"
#End If

   With RegEx

      ' clear all strings
      .Pattern = """[^""\r\n]*"""
      Code = .Replace(Code, "")

      ' remove comments
      '.Pattern = "'(.*)[\r\n]"
      .Pattern = "'(.*)(?:[\r\n]|$)"
      Code = .Replace(Code, "")

#If DebugPrintEnabled Then
   DebugPrint Code, True, "PrepareCode- after remove comments"
#End If

      ' dim a as String: a = 5  => insert line break
      .Pattern = "(\:\s)"
      Code = .Replace(Code, vbNewLine)

      ' remove Withevents
      .Pattern = "(Public|Private|Friend|Global)\sWithEvents\s"
      Code = .Replace(Code, "$1 ")

      ' API declaration => convert to normal procedure declaration
      .Pattern = "(?:Declare PtrSafe)\s(Function|Sub)\s"
      Code = .Replace(Code, "Declare $1 ")

      RegEx.Pattern = "(?:Declare)\s(Function|Sub)\s([^ ]*)[^(]*\("
      Code = .Replace(Code, "$1 $2(")

      ' remove Static before Function, Sub, Property, ..
      .Pattern = "(?:Static)+\s(Function|Sub|Property)\s"
      Code = .Replace(Code, "$1 ")

      ' remove Public, Private, Friend before Function, Sub, Property, ..
      .Pattern = "(?:Public|Private|Friend|Global)\s(Function|Sub|Property|Event|Enum|Type|Const)\s"
      Code = .Replace(Code, "$1 ")

   End With

#If DebugPrintEnabled Then
   DebugPrint Code, True, "PrepareCode - completed"
#End If

   PrepareCode = Code

End Function

Private Sub AddFromCode(ByVal Code As String, ByVal RegEx As RegExp, ByVal IsProcedure As Boolean, ByVal IsEnumTypeBlock As Boolean)

   Dim Match As Match
   Dim i As Long

   For Each Match In RegEx.Execute(Code)
      For i = 0 To Match.SubMatches.Count - 1
         AddWordFromDeclaration Match.SubMatches(i), IsProcedure, IsEnumTypeBlock
      Next
   Next

End Sub

Private Sub AddWordFromDeclaration(ByRef Declarations As String, ByVal IsProcedure As Boolean, ByVal IsEnumTypeBlock As Boolean)

   Dim Word As String
   Dim i As Long
   Dim Pos As Long
   Dim PosX As Long

   Dim DeclArray() As String

   If IsEnumTypeBlock Then
      Declarations = Replace(Declarations, vbCr, ",")
      Declarations = Replace(Declarations, vbLf, vbNullString)
   End If

   Declarations = Trim(Declarations)

   If IsProcedure Then
     ' Debug.Print Declarations
      Declarations = Replace(Declarations, "()", vbNullString)
      Declarations = Replace(Declarations, "Optional ", vbNullString)
      Declarations = Replace(Declarations, "ByRef ", vbNullString)
      Declarations = Replace(Declarations, "ByVal ", vbNullString)
      Declarations = Replace(Declarations, "ParamArray ", vbNullString)

      Pos = InStr(1, Declarations, "(")
      If Pos > 0 Then
         Mid(Declarations, Pos, 1) = ","
      End If
      Declarations = Replace(Declarations, ")", vbNullString)
   End If

   Do While InStr(1, Declarations, "  ") > 0
      Declarations = Replace(Declarations, "  ", " ")
   Loop

   DeclArray = Split(Trim(Declarations), ",")

   For i = LBound(DeclArray) To UBound(DeclArray)
      Word = Trim(DeclArray(i))
      Pos = CutPos(Word)
      If Pos > 1 Then
         Word = Trim(Left(Word, Pos - 1))
      End If
      If Len(Word) > 0 Then
         AddWord Word
      End If
   Next

End Sub

Private Function CutPos(ByVal Expression As String) As Long

   Dim Pos As Long
   Dim PosX As Long

   Const CutChars As String = " ()"

   Dim i As Long

   For i = 1 To Len(CutChars)
      PosX = InStr(1, Expression, Mid(CutChars, i, 1))
      If PosX > 0 Then
         If Pos = 0 Or PosX < Pos Then
            Pos = PosX
         End If
      End If
   Next

   CutPos = Pos

End Function

Public Function ToString(Optional ByVal ShowAll As Boolean = False) As String

   Dim WordKey As Variant
   Dim WordIndex As Long

   Dim VariationsDict As Scripting.Dictionary

   Dim OutputString As String

   For WordIndex = 0 To m_Words.Count - 1

      WordKey = m_Words.Keys(WordIndex)
      Set VariationsDict = m_Words.Item(WordKey)

      If VariationsDict.Count > (1 - Abs(ShowAll)) Then
         OutputString = OutputString & vbNewLine & WordKey & ":" & GetWordVariationsOutputString(WordKey)
      End If

   Next

   ToString = Mid(OutputString, Len(vbNewLine) + 1)

End Function

Public Function ToDict(Optional ByVal ShowAll As Boolean = False) As Scripting.Dictionary

   Dim WordKey As Variant
   Dim WordIndex As Long
   Dim OutputWord As Boolean
   Dim VariationsString As String

   Dim VariationsDict As Scripting.Dictionary

   Dim OutputDict As Scripting.Dictionary
   Set OutputDict = New Scripting.Dictionary

   OutputWord = ShowAll

   For WordIndex = 0 To m_Words.Count - 1

      WordKey = m_Words.Keys(WordIndex)
      Set VariationsDict = m_WordVariations.Item(WordKey)

      If Not ShowAll Then
         OutputWord = IsChangedItem(WordKey, VariationsDict)
      End If
      If OutputWord Then
         If VariationsDict.Count > 1 Then
            VariationsString = GetWordVariationsOutputString(WordKey)
         Else
            VariationsString = vbNullString
         End If
         OutputDict.Add WordKey, VariationsString
      End If

   Next

   Set ToDict = OutputDict

End Function

Private Function GetWordVariationsOutputString(ByVal Word As String, Optional ByVal IgnoreOriginalWord As Boolean = True) As String

   GetWordVariationsOutputString = Join(GetWordVariations(Word, IgnoreOriginalWord), "|")

End Function

Public Function GetWordVariations(ByVal Word As String, Optional ByVal IgnoreOriginalWord As Boolean = False) As String()

   Dim Variations() As String
   Dim VariantWord As String
   Dim AppendVariant As Boolean
   Dim i As Long, k As Long
   Dim VariationsDict As Scripting.Dictionary

   If StrComp(Word, m_Words.Item(Word), vbBinaryCompare) <> 0 Then
      Word = m_Words.Item(Word)
   End If

   Set VariationsDict = m_WordVariations.Item(Word)
   ReDim Variations(VariationsDict.Count - 1 - Abs(IgnoreOriginalWord))

   AppendVariant = Not IgnoreOriginalWord

   For i = 0 To VariationsDict.Count - 1
      VariantWord = VariationsDict.Keys(i)
      If IgnoreOriginalWord Then
         AppendVariant = StrComp(Word, VariantWord, vbBinaryCompare) <> 0
      End If
      If AppendVariant Then
         Variations(k) = VariantWord
         k = k + 1
      End If
   Next

   GetWordVariations = Variations

End Function

Public Sub ChangeKeyLetterCase(ByVal WordWithNewLetterCase As String)

   With m_Words
      .Key(WordWithNewLetterCase) = "~ " & WordWithNewLetterCase & " ~"
      .Key("~ " & WordWithNewLetterCase & " ~") = WordWithNewLetterCase
      .Item(WordWithNewLetterCase) = WordWithNewLetterCase
   End With

   m_DiffCnt = -1

End Sub

Public Sub LoadFromFile(ByVal FullFileName As String)

   With New ADODB.Stream
      .Charset = "utf-8"
      .Open
      .LoadFromFile FullFileName
      Do Until .EOS
          AddWord .ReadText(StreamReadEnum.adReadLine)
      Loop
      .Close
   End With

End Sub

Public Sub ExportToFile(ByVal FullFileName As String)

   Dim i As Long

   With New ADODB.Stream
      .Type = 2
      .Charset = "utf-8"
      .Open
      For i = 0 To m_Words.Count - 1
         .WriteText m_Words.Keys(i), adWriteLine
      Next
      .SaveToFile FullFileName, adSaveCreateOverWrite
      .Close
   End With

End Sub

#If DebugPrintEnabled Then

Private Sub DebugPrint(ByVal Text2Output As String, _
              Optional ByVal PrintHeaderLine As Boolean = False, _
              Optional ByVal HeaderLineCaption As String = vbNullString)

   Const HeaderLen As Long = 40
   Dim StringReplCount As Long
   If PrintHeaderLine Then
      StringReplCount = (HeaderLen - Len(HeaderLineCaption)) / 2
      If StringReplCount < 2 Then
         StringReplCount = 2
      End If
      Debug.Print String(StringReplCount, "-") & HeaderLineCaption & String(StringReplCount, "-")
   End If

   Debug.Print Text2Output

End Sub

#End If
