VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "uia_e"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==============================================================================
' uia_e (参照設定不要版)  ―  要素 (Element) のラッパ
'
' 参照アリ版との違い:
'   * elem / elemArray / tw は COM オブジェクトではなく生ポインタ (LongPtr)。
'     参照カウントは本クラスが管理する。
'   * New CUIAutomation を持たず、uia_core の共有 IUIAutomation を使う。
'   * ElementFromPoint は uia_core 側 (DispCallFunc) にまとめてある。
'
' この段階 (第1弾) で移植済み:
'   取得 / プロパティ / 検索(ff*/fa*) / TreeWalker / 基本パターン / 配列 / 座標
' 未移植 (第2弾予定): Edge*, IES_*(MSHTML), ptValue_DLFile, GetInfo/test 系
'==============================================================================
Option Explicit

#Const DebugMode = 1
#Const DebugMode2 = 0


' 保持する生ポインタ (参照カウントは本クラスが管理)
Private m_elem As LongPtr           ' IUIAutomationElement
Private m_elemArray As LongPtr      ' IUIAutomationElementArray
Private m_tw As LongPtr             ' IUIAutomationTreeWalker

Public TestMode As Boolean

Public Enum EnumTreeWakerType
    ControlView
    ContentView
    RawView
    ByCondition
End Enum

Public Enum EnumTreeWakeMoveTo
    MoveTo_Parent
    MoveTo_FirstChild
    MoveTo_LastChild
    MoveTo_Next
    MoveTo_Prev
End Enum

Public Enum EnumFindDirection
    Default = TreeTraversalOptions_Default
    PostOrder = TreeTraversalOptions_PostOrder
    LastToFirst = TreeTraversalOptions_LastToFirstOrder
End Enum

Public Enum EnumPattern
    Annotation = UIA_AnnotationPatternId
    CustomNavigation = UIA_CustomNavigationPatternId
    Dock = UIA_DockPatternId
    Drag = UIA_DragPatternId
    DropTarget = UIA_DropTargetPatternId
    ExpandCollapse = UIA_ExpandCollapsePatternId
    GridItem = UIA_GridItemPatternId
    Grid = UIA_GridPatternId
    Invoke = UIA_InvokePatternId
    ItemContainer = UIA_ItemContainerPatternId
    LegacyIAccessible = UIA_LegacyIAccessiblePatternId
    MultipleView = UIA_MultipleViewPatternId
    ObjectModel = UIA_ObjectModelPatternId
    RangeValue = UIA_RangeValuePatternId
    ScrollItem = UIA_ScrollItemPatternId
    Scroll = UIA_ScrollPatternId
    SelectionItem = UIA_SelectionItemPatternId
    Selection2 = UIA_SelectionPattern2Id
    Selection_ = UIA_SelectionPatternId
    SpreadsheetItem = UIA_SpreadsheetItemPatternId
    Spreadsheet = UIA_SpreadsheetPatternId
    Styles = UIA_StylesPatternId
    SynchronizedInput = UIA_SynchronizedInputPatternId
    TableItem = UIA_TableItemPatternId
    Table = UIA_TablePatternId
    TextChild = UIA_TextChildPatternId
    TextEdit = UIA_TextEditPatternId
    Text2 = UIA_TextPattern2Id
    Text = UIA_TextPatternId
    Toggle = UIA_TogglePatternId
    Transform2 = UIA_TransformPattern2Id
    Transform = UIA_TransformPatternId
    Value_ = UIA_ValuePatternId
    VirtualizedItem = UIA_VirtualizedItemPatternId
    Window = UIA_WindowPatternId
End Enum

'==============================================================================
' 保持ポインタ (elem / elemArray / tw)
'==============================================================================

Public Property Get elem() As LongPtr
    elem = m_elem
End Property
Public Property Let elem(ByVal p As LongPtr)
    If m_elem = p Then Exit Property
    If m_elem <> 0 Then uia_core.ComRelease m_elem
    m_elem = p
End Property

Public Property Get elemArray() As LongPtr
    elemArray = m_elemArray
End Property
Public Property Let elemArray(ByVal p As LongPtr)
    If m_elemArray = p Then Exit Property
    If m_elemArray <> 0 Then uia_core.ComRelease m_elemArray
    m_elemArray = p
End Property

Public Property Get tw() As LongPtr
    tw = m_tw
End Property
Public Property Let tw(ByVal p As LongPtr)
    If m_tw = p Then Exit Property
    If m_tw <> 0 Then uia_core.ComRelease m_tw
    m_tw = p
End Property

Private Sub Class_Terminate()
    If m_elem <> 0 Then uia_core.ComRelease m_elem
    If m_elemArray <> 0 Then uia_core.ComRelease m_elemArray
    If m_tw <> 0 Then uia_core.ComRelease m_tw
End Sub

'==============================================================================
' 起点エレメントの取得
'==============================================================================

Public Function getRoot() As uia_e
    Dim e_ As New uia_e
    e_.elem = uia_core.InvokeElem(uia_core.UIA(), 5, "GetRootElement")
    CopyMember e_
    Set getRoot = e_
End Function

Public Function getFocus() As uia_e
    Dim e_ As New uia_e
    e_.elem = uia_core.InvokeElem(uia_core.UIA(), 8, "GetFocusedElement")
    CopyMember e_
    Set getFocus = e_
End Function

Public Function getHandle(hWnd As LongPtr) As uia_e
    Dim e_ As New uia_e
    Dim p As LongPtr
    uia_core.UiaCheck uia_core.UiaInvoke(uia_core.UIA(), 6, "PP", hWnd, VarPtr(p)), _
                      "ElementFromHandle"
    e_.elem = p
    CopyMember e_
    Set getHandle = e_
End Function

Public Function getPoint(pt_ As tagPOINT) As uia_e
    Set getPoint = GetFromPoint(pt_.X, pt_.Y)
End Function

Public Function GetFromPoint(x As Long, y As Long) As uia_e
    Dim e_ As New uia_e
    e_.elem = uia_core.ElementFromPoint(x, y)
    CopyMember e_
    Set GetFromPoint = e_
End Function

Public Function GetFromCursor() As uia_e
    Dim pt As tagPOINT
    pt = uia_core.CursorPos()
    Set GetFromCursor = GetFromPoint(pt.X, pt.Y)
End Function

Public Sub Sleep_(dwMilliseconds As Long)
    uia_core.Sleep_ dwMilliseconds
End Sub

'==============================================================================
' エレメントの検索
'==============================================================================

Public Function ffChildren(c As uia_c, Optional RetryCount As Long = 0, Optional RetrySleepTime As Long = 300) As uia_e
    Dim e_ As New uia_e
    Set ffChildren = e_
    CopyMember e_
    If Alert_Nothing(m_elem, "ffChildren Root") Then Exit Function

    Dim cnt As Long, found As LongPtr
    For cnt = 0 To RetryCount
        found = 0
        uia_core.UiaCheck uia_core.UiaInvoke(m_elem, 5, "LPP", TreeScope_Children, c.cnd, VarPtr(found)), _
                          "IUIAutomationElement.FindFirst"
        e_.elem = found
        If Alert_Nothing(e_.elem, "ffChildren Target(" & cnt & ")") = False Then Exit For
        uia_core.Sleep_ RetrySleepTime
    Next
End Function

Public Function ffDescendants(c As uia_c, Optional RetryCount As Long = 0, Optional RetrySleepTime As Long = 300, Optional direction As EnumFindDirection = Default) As uia_e
    Dim e_ As New uia_e
    Set ffDescendants = e_
    CopyMember e_
    If Alert_Nothing(m_elem, "ffDescendants") Then Exit Function

    Dim cnt As Long, found As LongPtr
    For cnt = 0 To RetryCount
        found = FindFirstImpl(TreeScope_Descendants, c.cnd, direction)
        e_.elem = found
        If Alert_Nothing(e_.elem, "ffDescendants(" & cnt & ")") = False Then Exit For
        uia_core.Sleep_ RetrySleepTime
    Next
End Function

Public Function faChildren(Optional c As uia_c, Optional RetryCount As Long = 0, Optional RetrySleepTime As Long = 300) As uia_e
    If c Is Nothing Then
        Set c = New uia_c
        Set c = c.True_
    End If

    Dim e_ As New uia_e
    Set faChildren = e_
    CopyMember e_
    If Alert_Nothing(m_elem, "faChildren") Then Exit Function

    Dim cnt As Long, arr As LongPtr
    For cnt = 0 To RetryCount
        arr = 0
        uia_core.UiaCheck uia_core.UiaInvoke(m_elem, 6, "LPP", TreeScope_Children, c.cnd, VarPtr(arr)), _
                          "IUIAutomationElement.FindAll"
        e_.elemArray = arr
        If e_.Array_Length > 0 Then e_.elem = e_.ArrayElementPtr(0)
        If Alert_Nothing(e_.elem, "faChildren(" & cnt & ")") = False Then Exit For
        uia_core.Sleep_ RetrySleepTime
    Next
End Function

Public Function faDescendants(c As uia_c, Optional RetryCount As Long = 0, Optional RetrySleepTime As Long = 300, Optional direction As EnumFindDirection = Default) As uia_e
    Dim e_ As New uia_e
    Set faDescendants = e_
    CopyMember e_
    If Alert_Nothing(m_elem, "faDescendants") Then Exit Function

    Dim cnt As Long, arr As LongPtr
    For cnt = 0 To RetryCount
        arr = FindAllImpl(TreeScope_Descendants, c.cnd, direction)
        e_.elemArray = arr
        If e_.Array_Length > 0 Then e_.elem = e_.ArrayElementPtr(0)
        If arr <> 0 Then Exit For
        uia_core.Sleep_ RetrySleepTime
    Next
End Function

' direction=Default なら基底 FindFirst、そうでなければ Element9 の FindFirstWithOptions。
Private Function FindFirstImpl(ByVal scope As Long, ByVal cond As LongPtr, ByVal direction As Long) As LongPtr
    Dim found As LongPtr
    If direction = TreeTraversalOptions_Default Then
        uia_core.UiaCheck uia_core.UiaInvoke(m_elem, 5, "LPP", scope, cond, VarPtr(found)), _
                          "IUIAutomationElement.FindFirst"
    Else
        Dim p9 As LongPtr
        p9 = uia_core.ComQI(m_elem, IID_IUIAutomationElement9)
        If p9 = 0 Then
            uia_core.UiaCheck uia_core.UiaInvoke(m_elem, 5, "LPP", scope, cond, VarPtr(found)), _
                              "IUIAutomationElement.FindFirst"
        Else
            uia_core.UiaCheck uia_core.UiaInvoke(p9, 110, "LPLPP", scope, cond, direction, m_elem, VarPtr(found)), _
                              "IUIAutomationElement9.FindFirstWithOptions"
            uia_core.ComRelease p9
        End If
    End If
    FindFirstImpl = found
End Function

Private Function FindAllImpl(ByVal scope As Long, ByVal cond As LongPtr, ByVal direction As Long) As LongPtr
    Dim arr As LongPtr
    If direction = TreeTraversalOptions_Default Then
        uia_core.UiaCheck uia_core.UiaInvoke(m_elem, 6, "LPP", scope, cond, VarPtr(arr)), _
                          "IUIAutomationElement.FindAll"
    Else
        Dim p9 As LongPtr
        p9 = uia_core.ComQI(m_elem, IID_IUIAutomationElement9)
        If p9 = 0 Then
            uia_core.UiaCheck uia_core.UiaInvoke(m_elem, 6, "LPP", scope, cond, VarPtr(arr)), _
                              "IUIAutomationElement.FindAll"
        Else
            uia_core.UiaCheck uia_core.UiaInvoke(p9, 111, "LPLPP", scope, cond, direction, m_elem, VarPtr(arr)), _
                              "IUIAutomationElement9.FindAllWithOptions"
            uia_core.ComRelease p9
        End If
    End If
    FindAllImpl = arr
End Function

'==============================================================================
' TreeWalker
'==============================================================================

Public Sub tw_Set(twType As EnumTreeWakerType, Optional cnd As LongPtr)
    Dim p As LongPtr
    Select Case twType
        Case EnumTreeWakerType.ControlView
            uia_core.UiaCheck uia_core.UiaInvoke(uia_core.UIA(), 14, "P", VarPtr(p)), "get_ControlViewWalker"
        Case EnumTreeWakerType.ContentView
            uia_core.UiaCheck uia_core.UiaInvoke(uia_core.UIA(), 15, "P", VarPtr(p)), "get_ContentViewWalker"
        Case EnumTreeWakerType.RawView
            uia_core.UiaCheck uia_core.UiaInvoke(uia_core.UIA(), 16, "P", VarPtr(p)), "get_RawViewWalker"
        Case EnumTreeWakerType.ByCondition
            uia_core.UiaCheck uia_core.UiaInvoke(uia_core.UIA(), 13, "PP", cnd, VarPtr(p)), "CreateTreeWalker"
    End Select
    Me.tw = p
End Sub

Public Function twParent() As uia_e
    Set twParent = twMove(MoveTo_Parent)
End Function
Public Function twFirstChild() As uia_e
    Set twFirstChild = twMove(MoveTo_FirstChild)
End Function
Public Function twLastChild() As uia_e
    Set twLastChild = twMove(MoveTo_LastChild)
End Function
Public Function twNext() As uia_e
    Set twNext = twMove(MoveTo_Next)
End Function
Public Function twPrev() As uia_e
    Set twPrev = twMove(MoveTo_Prev)
End Function

Private Function twMove(MoveTo As EnumTreeWakeMoveTo) As uia_e
    Dim e_ As New uia_e
    Set twMove = e_
    If m_tw = 0 Then tw_Set ControlView
    CopyMember e_
    If Alert_Nothing(m_elem, "twMove Root") Then Exit Function

    Dim vtbl As Long
    Select Case MoveTo
        Case MoveTo_Parent:     vtbl = 3    ' GetParentElement
        Case MoveTo_FirstChild: vtbl = 4    ' GetFirstChildElement
        Case MoveTo_LastChild:  vtbl = 5    ' GetLastChildElement
        Case MoveTo_Next:       vtbl = 6    ' GetNextSiblingElement
        Case MoveTo_Prev:       vtbl = 7    ' GetPreviousSiblingElement
    End Select

    Dim p As LongPtr
    uia_core.UiaCheck uia_core.UiaInvoke(m_tw, vtbl, "PP", m_elem, VarPtr(p)), "TreeWalker.Move"
    e_.elem = p
    Alert_Nothing e_.elem, "twMove Target"
End Function

'==============================================================================
' プロパティ
'==============================================================================

Public Property Get prName() As String
    If Alert_Nothing(m_elem, "prName") Then Exit Property
    prName = uia_core.InvokeBstr(m_elem, 23, "get_CurrentName")
End Property

Public Property Get prClsName() As String
    If Alert_Nothing(m_elem, "prClsName") Then Exit Property
    prClsName = uia_core.InvokeBstr(m_elem, 30, "get_CurrentClassName")
End Property

Public Property Get prCtrlType() As Long
    If Alert_Nothing(m_elem, "prCtrlType") Then Exit Property
    prCtrlType = uia_core.InvokeLong(m_elem, 21, "get_CurrentControlType")
End Property

Public Property Get prLocalCtrlType() As String
    If Alert_Nothing(m_elem, "prLocalCtrlType") Then Exit Property
    prLocalCtrlType = uia_core.InvokeBstr(m_elem, 22, "get_CurrentLocalizedControlType")
End Property

Public Property Get prHwnd() As LongPtr
    If Alert_Nothing(m_elem, "prHwnd") Then Exit Property
    prHwnd = uia_core.GetPropertyValue(m_elem, UIA_NativeWindowHandlePropertyId)
End Property

Public Property Get prValue() As String
    If Alert_Nothing(m_elem, "prValue") Then Exit Property
    prValue = uia_core.GetPropertyValue(m_elem, UIA_ValueValuePropertyId)
End Property

Public Property Get prToggleState() As Long
    If Alert_Nothing(m_elem, "prToggleState") Then Exit Property
    prToggleState = uia_core.GetPropertyValue(m_elem, UIA_ToggleToggleStatePropertyId)
End Property

Public Property Get prExpandCollapseState() As Long
    If Alert_Nothing(m_elem, "prExpandCollapseState") Then Exit Property
    prExpandCollapseState = uia_core.GetPropertyValue(m_elem, UIA_ExpandCollapseExpandCollapseStatePropertyId)
End Property

Public Property Get prScrollState_X() As Double
    If Alert_Nothing(m_elem, "prScrollState_X") Then Exit Property
    prScrollState_X = uia_core.GetPropertyValue(m_elem, UIA_ScrollHorizontalScrollPercentPropertyId)
End Property

Public Property Get prScrollState_Y() As Double
    If Alert_Nothing(m_elem, "prScrollState_Y") Then Exit Property
    prScrollState_Y = uia_core.GetPropertyValue(m_elem, UIA_ScrollVerticalScrollPercentPropertyId)
End Property

Public Property Get prRect() As tagRECT
    If Alert_Nothing(m_elem, "prRect") Then Exit Property
    prRect = uia_core.GetBoundingRect(m_elem)
End Property

Public Property Get prRectLeft() As Long
    If Alert_Nothing(m_elem, "prRectLeft") Then Exit Property
    prRectLeft = uia_core.GetBoundingRect(m_elem).Left
End Property
Public Property Get prRectTop() As Long
    If Alert_Nothing(m_elem, "prRectTop") Then Exit Property
    prRectTop = uia_core.GetBoundingRect(m_elem).Top
End Property
Public Property Get prRectRight() As Long
    If Alert_Nothing(m_elem, "prRectRight") Then Exit Property
    prRectRight = uia_core.GetBoundingRect(m_elem).Right
End Property
Public Property Get prRectBottom() As Long
    If Alert_Nothing(m_elem, "prRectBottom") Then Exit Property
    prRectBottom = uia_core.GetBoundingRect(m_elem).Bottom
End Property

Public Property Get prRectCenter() As tagPOINT
    If Alert_Nothing(m_elem, "prRectCenter") Then Exit Property
    Dim r As tagRECT
    r = uia_core.GetBoundingRect(m_elem)
    prRectCenter.X = (r.Left + r.Right) / 2
    prRectCenter.Y = (r.Top + r.Bottom) / 2
End Property

'==============================================================================
' メソッド
'==============================================================================

Public Sub SetFocus()
    If Alert_Nothing(m_elem, "SetFocus") Then Exit Sub
    uia_core.UiaCheck uia_core.UiaInvoke(m_elem, 3, ""), "IUIAutomationElement.SetFocus"
End Sub

Public Sub ShowContextmenu()
    If Alert_Nothing(m_elem, "ShowContextMenu") Then Exit Sub
    Dim p3 As LongPtr
    p3 = uia_core.ComQI(m_elem, IID_IUIAutomationElement3)
    If p3 = 0 Then Exit Sub
    uia_core.UiaCheck uia_core.UiaInvoke(p3, 91, ""), "IUIAutomationElement3.ShowContextMenu"
    uia_core.ComRelease p3
End Sub

'==============================================================================
' パターン
'==============================================================================

Public Sub ptInvoke()
    Dim ptn As LongPtr
    ptn = GetPatternChecked("ptInvoke", Invoke)
    If ptn = 0 Then Exit Sub
    uia_core.UiaCheck uia_core.UiaInvoke(ptn, 3, ""), "InvokePattern.Invoke"
    uia_core.ComRelease ptn
End Sub

Public Sub ptSetValue(value As String)
    Dim ptn As LongPtr
    ptn = GetPatternChecked("ptSetValue", Value_)
    If ptn = 0 Then Exit Sub
    uia_core.UiaCheck uia_core.UiaInvoke(ptn, 3, "S", value), "ValuePattern.SetValue"
    uia_core.ComRelease ptn
End Sub

Public Sub ptSelectionItem_Select()
    Dim ptn As LongPtr
    ptn = GetPatternChecked("ptSelectionItem_Select", SelectionItem)
    If ptn = 0 Then Exit Sub
    uia_core.UiaCheck uia_core.UiaInvoke(ptn, 3, ""), "SelectionItemPattern.Select"
    uia_core.ComRelease ptn
End Sub

Public Function ptSelection_GetSelection() As uia_e
    Dim Target As New uia_e
    Set ptSelection_GetSelection = Target
    Dim ptn As LongPtr
    ptn = GetPatternChecked("ptSelection_GetSelection", Selection_)
    If ptn = 0 Then Exit Function
    Dim arr As LongPtr
    uia_core.UiaCheck uia_core.UiaInvoke(ptn, 3, "P", VarPtr(arr)), "SelectionPattern.GetCurrentSelection"
    uia_core.ComRelease ptn
    Target.elemArray = arr
    If Target.Array_Length > 0 Then Target.elem = Target.ArrayElementPtr(0)
End Function

Public Sub ptToggle()
    Dim ptn As LongPtr
    ptn = GetPatternChecked("ptToggle", Toggle)
    If ptn = 0 Then Exit Sub
    uia_core.UiaCheck uia_core.UiaInvoke(ptn, 3, ""), "TogglePattern.Toggle"
    uia_core.ComRelease ptn
End Sub

Public Sub ptExpand()
    Dim ptn As LongPtr
    ptn = GetPatternChecked("ptExpand", ExpandCollapse)
    If ptn = 0 Then Exit Sub
    uia_core.UiaCheck uia_core.UiaInvoke(ptn, 3, ""), "ExpandCollapsePattern.Expand"
    uia_core.ComRelease ptn
End Sub

Public Sub ptCollapse()
    Dim ptn As LongPtr
    ptn = GetPatternChecked("ptCollapse", ExpandCollapse)
    If ptn = 0 Then Exit Sub
    uia_core.UiaCheck uia_core.UiaInvoke(ptn, 4, ""), "ExpandCollapsePattern.Collapse"
    uia_core.ComRelease ptn
End Sub

Public Sub ptScroll(Optional XPercent As Double = 99999, Optional YPercent As Double = 99999)
    Dim ptn As LongPtr
    ptn = GetPatternChecked("ptScroll", Scroll)
    If ptn = 0 Then Exit Sub

    Dim sx As Double, sy As Double
    If XPercent = 99999 Then
        sx = InvokeDouble(ptn, 5, "ScrollPattern.get_CurrentHorizontalScrollPercent")
    Else
        sx = XPercent
    End If
    If YPercent = 99999 Then
        sy = InvokeDouble(ptn, 6, "ScrollPattern.get_CurrentVerticalScrollPercent")
    Else
        sy = YPercent
    End If
    uia_core.UiaCheck uia_core.UiaInvoke(ptn, 4, "DD", sx, sy), "ScrollPattern.SetScrollPercent"
    uia_core.ComRelease ptn
End Sub

Public Sub ptScrollIntoView()
    Dim ptn As LongPtr
    ptn = GetPatternChecked("ptScrollIntoView", ScrollItem)
    If ptn = 0 Then Exit Sub
    uia_core.UiaCheck uia_core.UiaInvoke(ptn, 3, ""), "ScrollItemPattern.ScrollIntoView"
    uia_core.ComRelease ptn
End Sub

Public Function ptGridGetItem(row As Long, col As Long) As uia_e
    Dim e_GridItem As New uia_e
    Set ptGridGetItem = e_GridItem
    Dim ptn As LongPtr
    ptn = GetPatternChecked("ptGridGetItem", Grid)
    If ptn = 0 Then Exit Function
    Dim p As LongPtr
    uia_core.UiaCheck uia_core.UiaInvoke(ptn, 3, "LLP", row, col, VarPtr(p)), "GridPattern.GetItem"
    uia_core.ComRelease ptn
    e_GridItem.elem = p
End Function

Public Sub ptWindowClose()
    Dim ptn As LongPtr
    ptn = GetPatternChecked("ptWindowClose", Window)
    If ptn = 0 Then Exit Sub
    uia_core.UiaCheck uia_core.UiaInvoke(ptn, 3, ""), "WindowPattern.Close"
    uia_core.ComRelease ptn
End Sub

Public Function ptGetTextRange() As uia_t
    Dim t_ As New uia_t
    Set ptGetTextRange = t_
    If Alert_Nothing(m_elem, "ptGetTextRange") Then Exit Function

    Dim ptn As LongPtr, rng As LongPtr
    Select Case True
        Case CBool(uia_core.GetPropertyValue(m_elem, UIA_IsTextPattern2AvailablePropertyId))
            ptn = uia_core.GetPattern(m_elem, UIA_TextPattern2Id)
            uia_core.UiaCheck uia_core.UiaInvoke(ptn, 7, "P", VarPtr(rng)), "TextPattern2.get_DocumentRange"
        Case CBool(uia_core.GetPropertyValue(m_elem, UIA_IsTextPatternAvailablePropertyId))
            ptn = uia_core.GetPattern(m_elem, UIA_TextPatternId)
            uia_core.UiaCheck uia_core.UiaInvoke(ptn, 7, "P", VarPtr(rng)), "TextPattern.get_DocumentRange"
        Case CBool(uia_core.GetPropertyValue(m_elem, UIA_IsTextChildPatternAvailablePropertyId))
            ptn = uia_core.GetPattern(m_elem, UIA_TextChildPatternId)
            uia_core.UiaCheck uia_core.UiaInvoke(ptn, 4, "P", VarPtr(rng)), "TextChildPattern.get_TextRange"
        Case CBool(uia_core.GetPropertyValue(m_elem, UIA_IsTextEditPatternAvailablePropertyId))
            ptn = uia_core.GetPattern(m_elem, UIA_TextEditPatternId)
            uia_core.UiaCheck uia_core.UiaInvoke(ptn, 7, "P", VarPtr(rng)), "TextEditPattern.get_DocumentRange"
        Case Else
            #If DebugMode = 1 Then
                Debug.Print "Any TextPattern Is Not Available! ptGetTextRange"
            #End If
            Exit Function
    End Select
    If ptn <> 0 Then uia_core.ComRelease ptn

    t_.tRng = rng
    t_.ScrollIntoView True
    t_.Select_
End Function

'==============================================================================
' elemArray を扱う処理
'==============================================================================

Public Sub Child_SetElemArray(Optional c As uia_c)
    Dim cnd_ As LongPtr, owns As Boolean
    If c Is Nothing Then
        cnd_ = uia_core.CreateTrueCondition()
        owns = True
    Else
        cnd_ = c.cnd
    End If
    Dim arr As LongPtr
    uia_core.UiaCheck uia_core.UiaInvoke(m_elem, 6, "LPP", TreeScope_Children, cnd_, VarPtr(arr)), _
                      "IUIAutomationElement.FindAll"
    Me.elemArray = arr
    If owns Then uia_core.ComRelease cnd_
End Sub

Public Function Child_GetItem(c As uia_c) As uia_e
    Dim i As Long, buf As uia_e, member As uia_e
    For i = 0 To Array_Length - 1
        Set member = Array_GetItemByIndex(i)
        Set buf = member.ffDescendants(c, 0)
        If Not Alert_Nothing(buf.elem) Then
            Set Child_GetItem = member
            Exit Function
        End If
    Next
End Function

Public Function Array_GetItemByIndex(Index As Long) As uia_e
    Dim ret_e As New uia_e
    Set Array_GetItemByIndex = ret_e
    If m_elemArray = 0 Then
        Debug.Print "elemArray Is Nothing! Array_GetItemByIndex(" & Index & ")"
        Exit Function
    End If
    If Array_Length - 1 < Index Then
        Debug.Print "elemArray MaxCount Over! Array_GetItemByIndex(" & Index & ")"
        Exit Function
    End If
    ret_e.elem = ArrayElementPtr(Index)
End Function

Public Property Get Array_Length() As Long
    If m_elemArray = 0 Then
        Debug.Print "elemArray Is Nothing! Array_Length"
        Exit Property
    End If
    Array_Length = uia_core.InvokeLong(m_elemArray, 3, "IUIAutomationElementArray.get_Length")
End Property

Public Property Get ElemArray_Col() As Collection
    Dim col As New Collection
    Set ElemArray_Col = col
    If m_elemArray = 0 Then
        Debug.Print "elemArray Is Nothing! ElemArrayCol"
        Exit Property
    End If
    Dim i As Long, ChildItem As uia_e
    For i = 0 To Array_Length - 1
        Set ChildItem = New uia_e
        CopyMember ChildItem
        ChildItem.elem = ArrayElementPtr(i)
        col.Add ChildItem
    Next
End Property

' 配列の要素ポインタ (AddRef 済み) を返す。GetElement = vtable[4]。
Public Function ArrayElementPtr(ByVal Index As Long) As LongPtr
    Dim p As LongPtr
    If m_elemArray = 0 Then Exit Function
    uia_core.UiaCheck uia_core.UiaInvoke(m_elemArray, 4, "LP", Index, VarPtr(p)), _
                      "IUIAutomationElementArray.GetElement"
    ArrayElementPtr = p
End Function

'==============================================================================
' 比較
'==============================================================================

Public Function CompareTo(Target As uia_e) As Boolean
    Dim same As Long
    uia_core.UiaCheck uia_core.UiaInvoke(uia_core.UIA(), 3, "PPP", m_elem, Target.elem, VarPtr(same)), _
                      "IUIAutomation.CompareElements"
    CompareTo = (same <> 0)
End Function

'==============================================================================
' 内部ヘルパ / デバッグ
'==============================================================================

' 引き継ぎ (tw と TestMode)。tw は共有するので AddRef する。
Private Sub CopyMember(e_ As uia_e)
    If m_tw <> 0 Then
        uia_core.ComAddRef m_tw
        e_.tw = m_tw
    End If
    e_.TestMode = TestMode
End Sub

' パターン取得 + 未サポート判定をまとめる。戻り値 0 なら未サポート。
Private Function GetPatternChecked(ByVal ctx As String, ByVal PtnId As EnumPattern) As LongPtr
    If Alert_Nothing(m_elem, ctx) Then Exit Function
    Dim p As LongPtr
    p = uia_core.GetPattern(m_elem, PtnId)
    If p = 0 Then
        #If DebugMode = 1 Then
            Debug.Print "Pattern Is Nothing! " & ctx & " (id=" & CLng(PtnId) & ")"
        #End If
    End If
    GetPatternChecked = p
End Function

Private Function InvokeDouble(ByVal pThis As LongPtr, ByVal vtblIndex As Long, ByVal ctx As String) As Double
    Dim v As Double
    uia_core.UiaCheck uia_core.UiaInvoke(pThis, vtblIndex, "P", VarPtr(v)), ctx
    InvokeDouble = v
End Function

Private Function Alert_Nothing(ByVal elemPtr As LongPtr, Optional FunctionName As String) As Boolean
    If TestMode Then Exit Function
    If elemPtr = 0 Then
        #If DebugMode = 1 Then
            Debug.Print "elem Is Nothing!(" & FunctionName & ")"
        #End If
        Alert_Nothing = True
    End If
End Function

Public Function GetCtrlTypeName(ctrlType As Long) As String
    Dim res As String
    Select Case ctrlType
        Case UIA_AppBarControlTypeId:       res = "UIA_AppBarControlTypeId"
        Case UIA_ButtonControlTypeId:       res = "UIA_ButtonControlTypeId"
        Case UIA_CalendarControlTypeId:     res = "UIA_CalendarControlTypeId"
        Case UIA_CheckBoxControlTypeId:     res = "UIA_CheckBoxControlTypeId"
        Case UIA_ComboBoxControlTypeId:     res = "UIA_ComboBoxControlTypeId"
        Case UIA_CustomControlTypeId:       res = "UIA_CustomControlTypeId"
        Case UIA_DataGridControlTypeId:     res = "UIA_DataGridControlTypeId"
        Case UIA_DataItemControlTypeId:     res = "UIA_DataItemControlTypeId"
        Case UIA_DocumentControlTypeId:     res = "UIA_DocumentControlTypeId"
        Case UIA_EditControlTypeId:         res = "UIA_EditControlTypeId"
        Case UIA_GroupControlTypeId:        res = "UIA_GroupControlTypeId"
        Case UIA_HeaderControlTypeId:       res = "UIA_HeaderControlTypeId"
        Case UIA_HeaderItemControlTypeId:   res = "UIA_HeaderItemControlTypeId"
        Case UIA_HyperlinkControlTypeId:    res = "UIA_HyperlinkControlTypeId"
        Case UIA_ImageControlTypeId:        res = "UIA_ImageControlTypeId"
        Case UIA_ListControlTypeId:         res = "UIA_ListControlTypeId"
        Case UIA_ListItemControlTypeId:     res = "UIA_ListItemControlTypeId"
        Case UIA_MenuBarControlTypeId:      res = "UIA_MenuBarControlTypeId"
        Case UIA_MenuControlTypeId:         res = "UIA_MenuControlTypeId"
        Case UIA_MenuItemControlTypeId:     res = "UIA_MenuItemControlTypeId"
        Case UIA_PaneControlTypeId:         res = "UIA_PaneControlTypeId"
        Case UIA_ProgressBarControlTypeId:  res = "UIA_ProgressBarControlTypeId"
        Case UIA_RadioButtonControlTypeId:  res = "UIA_RadioButtonControlTypeId"
        Case UIA_ScrollBarControlTypeId:    res = "UIA_ScrollBarControlTypeId"
        Case UIA_SemanticZoomControlTypeId: res = "UIA_SemanticZoomControlTypeId"
        Case UIA_SeparatorControlTypeId:    res = "UIA_SeparatorControlTypeId"
        Case UIA_SliderControlTypeId:       res = "UIA_SliderControlTypeId"
        Case UIA_SpinnerControlTypeId:      res = "UIA_SpinnerControlTypeId"
        Case UIA_SplitButtonControlTypeId:  res = "UIA_SplitButtonControlTypeId"
        Case UIA_StatusBarControlTypeId:    res = "UIA_StatusBarControlTypeId"
        Case UIA_TabControlTypeId:          res = "UIA_TabControlTypeId"
        Case UIA_TabItemControlTypeId:      res = "UIA_TabItemControlTypeId"
        Case UIA_TableControlTypeId:        res = "UIA_TableControlTypeId"
        Case UIA_TextControlTypeId:         res = "UIA_TextControlTypeId"
        Case UIA_ThumbControlTypeId:        res = "UIA_ThumbControlTypeId"
        Case UIA_TitleBarControlTypeId:     res = "UIA_TitleBarControlTypeId"
        Case UIA_ToolBarControlTypeId:      res = "UIA_ToolBarControlTypeId"
        Case UIA_ToolTipControlTypeId:      res = "UIA_ToolTipControlTypeId"
        Case UIA_TreeControlTypeId:         res = "UIA_TreeControlTypeId"
        Case UIA_TreeItemControlTypeId:     res = "UIA_TreeItemControlTypeId"
        Case UIA_WindowControlTypeId:       res = "UIA_WindowControlTypeId"
        Case Else:                          res = "UnKnownControlTypeId!"
    End Select
    GetCtrlTypeName = res
End Function

Public Property Get prShortCtrlTypeName() As String
    If Alert_Nothing(m_elem, "prShortCtrlTypeName") Then Exit Property
    Dim s As String
    s = GetCtrlTypeName(prCtrlType)
    s = Replace$(s, "UIA_", "")
    s = Replace$(s, "ControlTypeId", "")
    prShortCtrlTypeName = s
End Property
