VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeclarationDictTestCodemodule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'AccUnit:TestRelated

Dim AccUnitX As Long
Private m_AccUnitInfo As String

Private Enum TestEnum
   TestEnum_P1 = 2
End Enum
Private Type TypTest1
   FldA As Long
   FldB As String
End Type

Private Enum TestEnum2
   TestEnum2_P1 = 2
   TestEnum2_P2 = 3
End Enum

Private Type TypTest2
   Fld2A As Long
   Fld2B As String
End Type

Private Sub Class_Initialize()
'
End Sub

Public Function AccUnitTestFunct(ByVal FuncParam1 As Variant, FuncParam2() As String) As Variant

   Dim FuncVar1 As Variant
   Dim FuncVar2
   
   Dim FuncVar3()

End Function
