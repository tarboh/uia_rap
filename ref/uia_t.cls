VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "uia_t"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

#Const DebugMode = 1

Public tRng As IUIAutomationTextRange3

Public Enum EnumtRngSetType
    AsSelfInstance
    AsnewInstance
End Enum

Public Enum EnumTextUnitSize
    a_Character = TextUnit.TextUnit_Character
    b_Format = TextUnit.TextUnit_Format
    c_Word = TextUnit.TextUnit_Word
    d_Line = TextUnit.TextUnit_Line
    e_Paragraph = TextUnit.TextUnit_Paragraph
    f_Page = TextUnit.TextUnit_Page
    g_Document = TextUnit.TextUnit_Document
End Enum


Sub ScrollIntoView(AlignToTop As Boolean)
    If Alert_TextRangeAvailable(tRng, "ScrollIntoView") Then Exit Sub
    Call tRng.ScrollIntoView(AlignToTop)
End Sub

Sub Select_()
    If Alert_TextRangeAvailable(tRng, "Select_") Then Exit Sub
    tRng.Select
End Sub
    

Function Find(Returntype As EnumtRngSetType, val As String, Optional FindFromLast As Boolean, Optional IgnoreCase As Boolean) As uia_t
    Dim Target As uia_t
    If Alert_TextRangeAvailable(tRng, "Find") Then Exit Function
    If Returntype = AsnewInstance Then
        Set Target = New uia_t
        Set Target.tRng = tRng.FindText(val, FindFromLast, IgnoreCase)
    Else
        Set Target = Me
        Set Target.tRng = tRng.FindText(val, FindFromLast, IgnoreCase)
    End If
    Set Find = Target
    If Alert_TextRangeAvailable(Target.tRng, "Find") Then Exit Function
    Target.tRng.ScrollIntoView True
    Target.tRng.Select
End Function

Function Expand(Returntype As EnumtRngSetType, UnitSize As EnumTextUnitSize) As uia_t
    Dim Target As uia_t
    If Alert_TextRangeAvailable(tRng, "Expand") Then Exit Function
    If Returntype = AsnewInstance Then
        Set Target = New uia_t
        Set Target.tRng = Me.tRng
        Call Target.tRng.ExpandToEnclosingUnit(UnitSize)
    Else
        Set Target = Me
        Call Target.tRng.ExpandToEnclosingUnit(UnitSize)
    End If
    Set Expand = Target
    If Alert_TextRangeAvailable(Target.tRng, "Expand") Then Exit Function
    Target.tRng.ScrollIntoView True
    Target.tRng.Select
End Function

Function Move(Returntype As EnumtRngSetType, UnitSize As EnumTextUnitSize, MoveSize As Long) As uia_t
    Dim Target As uia_t
    If Alert_TextRangeAvailable(tRng, "Move") Then Exit Function
    If Returntype = AsnewInstance Then
        Set Target = New uia_t
        Set Target.tRng = Me.tRng
    Else
        Set Target = Me
    End If
    Call tRng.MoveEndpointByUnit(TextPatternRangeEndpoint_End, UnitSize, MoveSize)
    Call tRng.MoveEndpointByUnit(TextPatternRangeEndpoint_Start, UnitSize, MoveSize)
    tRng.Select
    tRng.ScrollIntoView True

End Function


'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'デバッグ用

Function Alert_InstanceIsAvailable(c As uia_t, Optional Name As String) As Boolean
    If c Is Nothing Then
        #If DebugMode = 1 Then
            Debug.Print "uia_t Is Nothig! " & Name
        #End If
        Alert_InstanceIsAvailable = True
    End If
End Function

Function Alert_TextRangeAvailable(tRng As IUIAutomationTextRange, Name As String) As Boolean
    If tRng Is Nothing Then
        #If DebugMode = 1 Then
            Debug.Print "tRng Is Nothig! " & Name
        #End If
        Alert_TextRangeAvailable = True
    End If
End Function
