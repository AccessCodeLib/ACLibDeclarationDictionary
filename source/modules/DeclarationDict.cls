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

Private m_Words As Scripting.Dictionary
Private m_DiffCnt As Long

Private Sub Class_Initialize()
   Set m_Words = GetNewDict(TextCompare)
   m_DiffCnt = -1
End Sub

Private Function GetNewDict(ByVal CompareMode As Scripting.CompareMethod) As Scripting.Dictionary
   Dim Dict As Scripting.Dictionary
   Set Dict = New Scripting.Dictionary
   Dict.CompareMode = CompareMode
   Set GetNewDict = Dict
End Function

Friend Property Get Dict() As Scripting.Dictionary
   Set Dict = m_Words
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
      If IsChangedItem(Word, m_Words.Item(Word)) Then
         DifCntr = DifCntr + 1
      End If
   Next

   GetDiffCnt = DifCntr

End Function

Private Function IsChangedItem(ByVal Word As String, ByVal SubDict As Scripting.Dictionary) As Boolean

   Dim LastKey As String

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

   m_DiffCnt = -1

   Dim SubDict As Scripting.Dictionary
   Dim LastKey As String

   If Not m_Words.Exists(Word) Then
      m_Words.Add Word, GetNewDict(BinaryCompare)
   End If

   Set SubDict = m_Words.Item(Word)
   With SubDict
      If Not .Exists(Word) Then
         .Add Word, Word
      ElseIf .Count > 1 Then
         LastKey = .Keys(.Count - 1)
         If StrComp(LastKey, Word, vbBinaryCompare) <> 0 Then
            .Remove Word
            .Add Word, Word
         End If
      End If
   End With

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
   Set RegEx = New RegExp

   Dim i As Long
   Dim IsProcedure As Boolean

   Code = Replace(Code, " _" & vbNewLine, " ")

   RegEx.IgnoreCase = True
   RegEx.Global = True

   RegEx.Pattern = """[^""\r\n]*"""
   Code = RegEx.Replace(Code, "")

   RegEx.Pattern = "'(.*)[\r\n]"
   Code = RegEx.Replace(Code, "")

   RegEx.Pattern = "(?:Declare PtrSafe)\s(Function|Sub)\s"
   Code = RegEx.Replace(Code, "Declare $1 ")

   RegEx.Pattern = "(Declare)\s(Function|Sub)\s(.*)Alias\s*[(]"
   Code = RegEx.Replace(Code, "$1 $2 $3(")

   RegEx.Pattern = "(?:Declare)\s(Function|Sub)\s(.*)Lib\s*[(]"
   Code = RegEx.Replace(Code, "$1 $2(")

   RegEx.Pattern = "(?:Public|Private|Friend)\s(Function|Sub|Property|Enum|Const)\s"
   Code = RegEx.Replace(Code, "$1 ")

   Dim Patterns(1) As String
   Patterns(0) = "(?:\r|\n)\s*(?:Sub|Function|Property Get|Property Let)\s+(.*)\s*"
   Patterns(1) = "(?:\r|\n)\s*(?:Dim|Private|Friend|Public|Const|Enum)\s+(.*)"

   For i = 0 To UBound(Patterns)
      RegEx.Pattern = Patterns(i)
      AddFromCode Code, RegEx, i = 0
   Next

End Sub

Private Sub AddFromCode(ByVal Code As String, ByVal RegEx As RegExp, ByVal IsProcedure As Boolean)

   Dim Match As Match
   Dim i As Long

   For Each Match In RegEx.Execute(Code)
      For i = 0 To Match.SubMatches.Count - 1
         AddWordFromDeclaration Match.SubMatches(i), IsProcedure
      Next
   Next

End Sub

Private Sub AddWordFromDeclaration(ByRef Declarations As String, ByVal IsProcedure As Boolean)

   Dim Word As String
   Dim i As Long
   Dim Pos As Long
   Dim PosX As Long

   'Stop
   'Debug.Print Declarations, IsProcedure

   Dim DeclArray() As String

   Declarations = Replace(Declarations, vbCr, vbNullString)
   Declarations = Replace(Declarations, vbLf, vbNullString)
   Declarations = Trim(Declarations)

   If IsProcedure Then
      Declarations = Replace(Declarations, "()", vbNullString)
      Declarations = Replace(Declarations, "Optional ", vbNullString)
      Declarations = Replace(Declarations, "ByRef ", vbNullString)
      Declarations = Replace(Declarations, "ByVal ", vbNullString)

      Pos = InStr(1, Declarations, "(")
      If Pos > 0 Then
         Mid(Declarations, Pos, 1) = ","
      End If
      Declarations = Replace(Declarations, ")", vbNullString)
   End If

   DeclArray = Split(Trim(Declarations), ",")

   For i = LBound(DeclArray) To UBound(DeclArray)
      Word = Trim(DeclArray(i))
      Pos = CutPos(Word)
      If Pos > 1 Then
         Word = Left(Word, Pos - 1)
      End If
      AddWord Word
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

   Dim VariantsDict As Scripting.Dictionary

   Dim OutputString As String

   For WordIndex = 0 To m_Words.Count - 1

      WordKey = m_Words.Keys(WordIndex)
      Set VariantsDict = m_Words.Item(WordKey)

      If VariantsDict.Count > (1 - Abs(ShowAll)) Then
         OutputString = OutputString & vbNewLine & WordKey & ":" & GetWordVariantsOutputString(WordKey)
      End If

   Next

   ToString = Mid(OutputString, Len(vbNewLine) + 1)

End Function

Public Function ToDict(Optional ByVal ShowAll As Boolean = False) As Scripting.Dictionary

   Dim WordKey As Variant
   Dim WordIndex As Long
   Dim OutputWord As Boolean
   Dim VariantsString As String

   Dim VariantsDict As Scripting.Dictionary

   Dim OutputDict As Scripting.Dictionary
   Set OutputDict = New Scripting.Dictionary

   OutputWord = ShowAll

   For WordIndex = 0 To m_Words.Count - 1

      WordKey = m_Words.Keys(WordIndex)
      Set VariantsDict = m_Words.Item(WordKey)

      If Not ShowAll Then
         OutputWord = IsChangedItem(WordKey, VariantsDict)
      End If
      If OutputWord Then
         If VariantsDict.Count > 1 Then
            VariantsString = GetWordVariantsOutputString(WordKey)
         Else
            VariantsString = vbNullString
         End If
         OutputDict.Add WordKey, VariantsString
      End If

   Next

   Set ToDict = OutputDict

End Function

Private Function GetWordVariantsOutputString(ByVal Word As String, Optional ByVal IgnoreOriginalWord As Boolean = True) As String

   GetWordVariantsOutputString = Join(GetWordVariants(Word, IgnoreOriginalWord), "|")

End Function


Public Function GetWordVariants(ByVal Word As String, Optional ByVal IgnoreOriginalWord As Boolean = False) As String()

   Dim Variants() As String
   Dim VariantWord As String
   Dim AppendVariant As Boolean
   Dim i As Long, k As Long
   Dim VariantsDict As Scripting.Dictionary

   Set VariantsDict = m_Words.Item(Word)
   ReDim Variants(VariantsDict.Count - 1 - Abs(IgnoreOriginalWord))

   AppendVariant = Not IgnoreOriginalWord

   For i = 0 To VariantsDict.Count - 1
      VariantWord = VariantsDict.Keys(i)
      If IgnoreOriginalWord Then
         AppendVariant = StrComp(Word, VariantWord, vbBinaryCompare) <> 0
      End If
      If AppendVariant Then
         Variants(k) = VariantWord
         k = k + 1
      End If
   Next

   GetWordVariants = Variants

End Function

Public Sub ChangeKeyLetterCase(ByVal WordWithNewLetterCase As String)

   With m_Words
      .Key(WordWithNewLetterCase) = "~ " & WordWithNewLetterCase & " ~"
      .Key("~ " & WordWithNewLetterCase & " ~") = WordWithNewLetterCase
   End With

   m_DiffCnt = -1

End Sub
