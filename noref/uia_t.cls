VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "uia_t"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==============================================================================
' uia_t (参照設定不要版)  ―  テキスト範囲 (TextRange) のラッパ
'
' 参照アリ版との違い:
'   * tRng は IUIAutomationTextRange3 オブジェクトではなく、生の範囲ポインタ
'     (LongPtr)。参照カウントは本クラスが管理する。
'   * 各操作は uia_core.UiaInvoke (DispCallFunc) で TextRange の vtable を叩く。
'==============================================================================
Option Explicit

#Const DebugMode = 1

' テキスト範囲ポインタ (IUIAutomationTextRange)。参照カウントは本クラスが持つ。
Private m_tRng As LongPtr

Public Enum EnumtRngSetType
    AsSelfInstance
    AsnewInstance
End Enum

Public Enum EnumTextUnitSize
    a_Character = TextUnit_Character
    b_Format = TextUnit_Format
    c_Word = TextUnit_Word
    d_Line = TextUnit_Line
    e_Paragraph = TextUnit_Paragraph
    f_Page = TextUnit_Page
    g_Document = TextUnit_Document
End Enum

'==============================================================================
' tRng (生の範囲ポインタ)
'==============================================================================

Public Property Get tRng() As LongPtr
    tRng = m_tRng
End Property

' 範囲ポインタを設定する。所有権を引き受け、古い範囲は解放する。
' 同じ範囲を共有したい場合は呼び出し側で uia_core.ComAddRef してから渡すこと。
Public Property Let tRng(ByVal newPtr As LongPtr)
    If m_tRng = newPtr Then Exit Property
    If m_tRng <> 0 Then uia_core.ComRelease m_tRng
    m_tRng = newPtr
End Property

Private Sub Class_Terminate()
    If m_tRng <> 0 Then uia_core.ComRelease m_tRng
End Sub

Private Function Avail(ByVal ctx As String) As Boolean
    If m_tRng = 0 Then
        #If DebugMode = 1 Then
            Debug.Print "tRng Is Null! " & ctx
        #End If
        Exit Function
    End If
    Avail = True
End Function

'==============================================================================
' 操作
'==============================================================================

Public Sub ScrollIntoView(ByVal AlignToTop As Boolean)
    If Not Avail("ScrollIntoView") Then Exit Sub
    uia_core.UiaCheck uia_core.UiaInvoke(m_tRng, 19, "L", IIf(AlignToTop, 1, 0)), _
                      "TextRange.ScrollIntoView"
End Sub

Public Sub Select_()
    If Not Avail("Select_") Then Exit Sub
    uia_core.UiaCheck uia_core.UiaInvoke(m_tRng, 16, ""), "TextRange.Select"
End Sub

' 範囲のテキストを取得する。maxLength = -1 で全文。
Public Function GetText(Optional ByVal maxLength As Long = -1) As String
    If Not Avail("GetText") Then Exit Function
    uia_core.UiaCheck uia_core.UiaInvoke(m_tRng, 12, "LP", maxLength, VarPtr(GetText)), _
                      "TextRange.GetText"
End Function

Public Function Find(Returntype As EnumtRngSetType, val_ As String, Optional FindFromLast As Boolean, Optional IgnoreCase As Boolean) As uia_t
    If Not Avail("Find") Then Exit Function

    Dim found As LongPtr
    uia_core.UiaCheck uia_core.UiaInvoke(m_tRng, 8, "SLLP", val_, _
        IIf(FindFromLast, 1, 0), IIf(IgnoreCase, 1, 0), VarPtr(found)), "TextRange.FindText"

    Dim Target As uia_t
    If Returntype = AsnewInstance Then Set Target = New uia_t Else Set Target = Me
    Target.tRng = found                 ' 所有権移譲

    Set Find = Target
    If Target.tRng = 0 Then Exit Function
    Target.ScrollIntoView True
    Target.Select_
End Function

Public Function Expand(Returntype As EnumtRngSetType, UnitSize As EnumTextUnitSize) As uia_t
    If Not Avail("Expand") Then Exit Function

    Dim Target As uia_t
    If Returntype = AsnewInstance Then
        Set Target = New uia_t
        uia_core.ComAddRef m_tRng       ' 同じ範囲を共有 (原本の挙動)
        Target.tRng = m_tRng
    Else
        Set Target = Me
    End If

    uia_core.UiaCheck uia_core.UiaInvoke(Target.tRng, 6, "L", CLng(UnitSize)), _
                      "TextRange.ExpandToEnclosingUnit"

    Set Expand = Target
    If Target.tRng = 0 Then Exit Function
    Target.ScrollIntoView True
    Target.Select_
End Function

Public Function Move(Returntype As EnumtRngSetType, UnitSize As EnumTextUnitSize, MoveSize As Long) As uia_t
    If Not Avail("Move") Then Exit Function

    Dim Target As uia_t
    If Returntype = AsnewInstance Then
        Set Target = New uia_t
        uia_core.ComAddRef m_tRng
        Target.tRng = m_tRng
    Else
        Set Target = Me
    End If

    ' 原本どおり Me 側の範囲に対して操作する
    Dim moved As Long
    uia_core.UiaInvoke m_tRng, 14, "LLLP", TextPatternRangeEndpoint_End, CLng(UnitSize), MoveSize, VarPtr(moved)
    uia_core.UiaInvoke m_tRng, 14, "LLLP", TextPatternRangeEndpoint_Start, CLng(UnitSize), MoveSize, VarPtr(moved)
    uia_core.UiaInvoke m_tRng, 16, ""            ' Select
    uia_core.UiaInvoke m_tRng, 19, "L", 1        ' ScrollIntoView True

    Set Move = Target                            ' 原本は戻り値未設定だったので補った
End Function

'==============================================================================
' デバッグ
'==============================================================================

Function Alert_InstanceIsAvailable(c As uia_t, Optional Name As String) As Boolean
    If c Is Nothing Then
        #If DebugMode = 1 Then
            Debug.Print "uia_t Is Nothing! " & Name
        #End If
        Alert_InstanceIsAvailable = True
    End If
End Function
