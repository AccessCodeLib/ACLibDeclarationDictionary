VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeclarationDictTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestClass

Private m_DeclDict As DeclarationDict

'--------------------------------------------------------------------
' Test Preparation / Cleanup
'--------------------------------------------------------------------
Public Sub Setup()
   Set m_DeclDict = New DeclarationDict
End Sub

Public Sub TearDown()
   Set m_DeclDict = Nothing
End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

'AccUnit:Row("Public Function Func1() As String", "Func1")
'AccUnit:Row("Private Function Func1() As String", "Func1")
'AccUnit:Row("Private Function Func1() As String()", "Func1")
'AccUnit:Row("Function Func1()", "Func1")
'AccUnit:Row("Dim Abc As String", "Abc")
'AccUnit:Row("Dim Abc", "Abc")
'AccUnit:Row("Dim Abc() As String", "Abc")
'AccUnit:Row("Dim Abc()", "Abc")
'AccUnit:Row("Private Abc As String", "Abc")
'AccUnit:Row("Private Abc", "Abc")
'AccUnit:Row("Private Abc() As String", "Abc")
'AccUnit:Row("Private Abc()", "Abc")
'AccUnit:Row("Private Function Func1(a, b) As String()", "Func1|a|b")
'AccUnit:Row("Private Function Func1(a(), b) As String()", "Func1|a|b")
'AccUnit:Row("Private Function Func1(a(), b()) As String()", "Func1|a|b")
'AccUnit:Row("Dim Abc() As String, b(), C As Long", "Abc|b|C")
'AccUnit:Row("Function Abc(ByVal X as Long) as Long: Abc = X*X: End Function", "Abc|X")
Public Sub ImportCode_InsertCodeLine_CheckKeysExists(ByVal Code As String, ByVal Expected As String)
  
   Dim Actual As String
      
   m_DeclDict.ImportCode Code
   Actual = Join(m_DeclDict.Dict.Keys, "|")
   
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub

'AccUnit:Row("Private Function Func1() As String" & Environment.NewLine & Environment.NewLine & "   Dim X as String, Y As Long", "Func1|X|Y")
'AccUnit:Row(Environment.NewLine & "Private abc As String" & Environment.NewLine & Environment.NewLine & "   Public X as String, Y As Long", "abc|X|Y")
Public Sub ImportCode_InsertCodeLines_CheckKeysExists(ByVal Code As String, ByVal Expected As String)

   Dim Actual As String
   
   Debug.Print Code
      
   m_DeclDict.ImportCode Code
   Actual = Join(m_DeclDict.Dict.Keys, "|")
   
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub

Public Sub ImportCode_EnumCode_CheckKeys()

   Dim Actual As String
   Dim Code As String
   
   Code = "Dim x as long" & vbNewLine & _
          "Private Enum xyz" & vbNewLine & _
          "   en1 = 0" & vbNewLine & _
          "   en2 = 1" & vbNewLine & _
          "End Enum" & vbNewLine & _
          "Dim y()"
   
   Const Expected As String = "xyz|en1|en2|x|y"

   m_DeclDict.ImportCode Code
   Actual = Join(m_DeclDict.Dict.Keys, "|")
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub ImportCode_TypeCode_CheckKeys()

   Dim Actual As String
   Dim Code As String
   
   Code = "Dim x as long" & vbNewLine & _
          "Private Type xyz" & vbNewLine & _
          "   F1 As String" & vbNewLine & _
          "   F2 As Long" & vbNewLine & _
          "End Type" & vbNewLine & _
          "Dim y()"
   
   Const Expected As String = "xyz|F1|F2|x|y"

   m_DeclDict.ImportCode Code
   Actual = Join(m_DeclDict.Dict.Keys, "|")
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub ImportCodeModule_CheckKeys()

   Dim ActualDict As Scripting.Dictionary
   
   Dim ExpectedKeys() As Variant
   ExpectedKeys = Array("AccUnitX", "m_AccUnitInfo", "Field", _
                        "TestEnum", "TestEnum_P1", _
                        "TestType1", "FldA", "FldB", "FldC", _
                        "TestEnum2", "TestEnum2_P1", "TestEnum2_P2", _
                        "TestType2", "Fld2A", "Fld2B", _
                        "RaiseSomething", "EventParam1", _
                        "Sleep", "dwMilliseconds", _
                        "RaiseSomething2", "EventParam2", _
                        "Class_Initialize", _
                        "AccUnitTestFunct", "FuncParam1", "FuncParam2", "FuncVar1", "FuncVar2", "FuncVar3", "FuncVar4", "Dim1", "Dim2", _
                        "Name1", "Name2", "NewValue", _
                        "PropertySet", "ObjRef", _
                        "TestMe")
                        
   '
   m_DeclDict.ImportVBComponent CurrentVbProject.VBComponents("DeclarationDictTestCodemodule")
   
   Set ActualDict = m_DeclDict.Dict
   
   Assert.That ActualDict.Count, Iz.EqualTo(UBound(ExpectedKeys) + 1), "Count of items"
   
   
   Dim i As Long
   For i = 0 To UBound(ExpectedKeys)
      Assert.IsTrue ActualDict.Exists(ExpectedKeys(i)), ExpectedKeys(i) & " not exists"
   Next
   
End Sub

