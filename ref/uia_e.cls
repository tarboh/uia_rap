VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "uia_e"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

#Const DebugMode = 0
#Const DebugMode2 = 0

Public TestMode As Boolean

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

#If VBA7 Then
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
        ByVal pCaller As LongPtr, _
        ByVal szURL As String, _
        ByVal szFileName As String, _
        ByVal dwReserved As Long, _
        ByVal lpfnCB As LongPtr) As Long
#Else
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
        ByVal pCaller As Long, _
        ByVal szURL As String, _
        ByVal szFileName As String, _
        ByVal dwReserved As Long, _
        ByVal lpfnCB As Long) As Long
#End If

Private uia As New CUIAutomation
Public elem As IUIAutomationElement9
Public elemArray As IUIAutomationElementArray

Public tw As IUIAutomationTreeWalker

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

Private Type PointAPI
    x As Long
    y As Long
End Type

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

#If VBA7 Then
    Private Declare PtrSafe Function GetCursorPos Lib "user32.dll" (lpPoint As PointAPI) As Long
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    #If Win64 Then
        Private Declare PtrSafe Function AccessibleObjectFromPoint Lib "oleacc" (ByVal arg1 As LongPtr, ppacc As IAccessible, pvarChild As Variant) As Long
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    #Else
        Private Declare PtrSafe Function AccessibleObjectFromPoint Lib "oleacc" (ByVal lX As Long, ByVal lY As Long, ppacc As IAccessible, pvarChild As Variant) As Long
    #End If
#Else
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
    Public Declare Function AccessibleObjectFromPoint Lib "Oleacc" (ByVal lX As Long, ByVal lY As Long, ppacc As IAccessible, pvarChild As Variant) As Long
    Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If

'ライブラリ内の関数を手動で呼び出すAPI
    Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" ( _
        ByVal pvInstance As LongPtr, _
        ByVal offsetinVft As Long, _
        ByVal CallConv As Long, _
        ByVal retTYP As Integer, _
        ByVal paCNT As Long, _
        ByRef paTypes As Integer, _
        ByRef paValues As LongPtr, _
        ByRef retVAR As Variant _
    ) As Long

    Private Const S_OK = 0
    Private Const CC_STDCALL As Long = 4

#If Win64 Then
    Private Const pIUnknownRelease As Long = 8   'Address of 2nd Function in the virtual function table of the IUIAutomationElement Interface : (2-1)th * 8 Byte = "8"
    Private Const pElementFromPoint As Long = 56 'Address of 8th Function in the virtual function table of the CUIAutomation class            : (8-1)th * 8 Byte = "56"
    Private Const pCount = 2                     'Count of arguments in 64bit environment ( pt, Element* )
#Else
    Private Const pIUnknownRelease As Long = 4   'Address of 2nd Function in the virtual function table of the IUIAutomationElement Interface : (2-1)th * 4 Byte = "4"
    Private Const pElementFromPoint As Long = 28 'Address of 8th Function in the virtual function table of the CUIAutomation class            : (8-1)th * 4 Byte = "28"
    Private Const pCount = 3                     'Count of arguments in 32bit environment ( pt.x, pt.y, Element* )
#End If

'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'起点エレメントの取得
Public Function getRoot() As uia_e
    Dim e_ As New uia_e
    Set e_.elem = uia.GetRootElement
    Call CopyMember(e_)
    Set getRoot = e_
End Function

Public Function getFocus() As uia_e
    Dim e_ As New uia_e
    Set e_.elem = uia.GetFocusedElement
    Call CopyMember(e_)
    Set getFocus = e_
End Function

Public Function getHandle(hWnd As LongPtr) As uia_e
    Dim e_ As New uia_e
    Set e_.elem = uia.ElementFromHandle(ByVal hWnd)
    Call CopyMember(e_)
    Set getHandle = e_
End Function

Public Function getPoint(pt_ As tagPOINT) As uia_e
    
End Function

Public Sub Sleep_(dwMilliseconds As Long)
    Call Sleep(dwMilliseconds)
End Sub

'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'エレメントの検索
Public Function ffChildren(c As uia_c, Optional RetryCount As Long = 0, Optional RetrySleepTime As Long = 300) As uia_e
    Dim e_ As New uia_e
    Set ffChildren = e_
    Call CopyMember(e_)
    If Alert_Nothing(elem, "ff_Children Root") Then Exit Function
    
    '取得に失敗した場合は一定回数繰り返す
    Dim cnt As Long
    For cnt = 0 To RetryCount
        Set e_.elem = elem.FindFirst(TreeScope_Children, c.cnd)
        If Alert_Nothing(e_.elem, "ff_Children Target(" & cnt & ")") = False Then Exit For
        Sleep RetrySleepTime
    Next
End Function

Public Function ffDescendants(c As uia_c, Optional RetryCount As Long = 0, Optional RetrySleepTime As Long = 300, Optional direction As EnumFindDirection = Default) As uia_e
    Dim e_ As New uia_e
    Set ffDescendants = e_
    Call CopyMember(e_)
    If Alert_Nothing(elem, "ff_Descendants") Then Exit Function
    
    '取得に失敗した場合は一定回数繰り返す
    Dim cnt As Long
    For cnt = 0 To RetryCount
        Set e_.elem = elem.FindFirstWithOptions(TreeScope_Descendants, c.cnd, direction, elem)
        If Alert_Nothing(e_.elem, "ff_Descendants(" & cnt & ")") = False Then Exit For
        Sleep RetrySleepTime
    Next
End Function

Public Function faChildren(Optional c As uia_c, Optional RetryCount As Long = 0, Optional RetrySleepTime As Long = 300) As uia_e
    
    If c Is Nothing Then
        Set c = New uia_c
        Set c = c.True_
    End If
    
    Dim e_ As New uia_e
    Set faChildren = e_
    Call CopyMember(e_)
    If Alert_Nothing(elem, "fa_Children") Then Exit Function
    
    '取得に失敗した場合は一定回数繰り返す
    Dim cnt As Long
    For cnt = 0 To RetryCount
        Set e_.elemArray = elem.FindAll(TreeScope_Children, c.cnd)
        If e_.elemArray.Length > 0 Then Set e_.elem = e_.elemArray.GetElement(0)
        If Alert_Nothing(e_.elem, "fa_Children(" & cnt & ")") = False Then Exit For
        Sleep RetrySleepTime
    Next
    
End Function

Public Function faDescendants(c As uia_c, Optional RetryCount As Long = 0, Optional RetrySleepTime As Long = 300, Optional direction As EnumFindDirection = Default) As uia_e
    Dim e_ As New uia_e
    Set faDescendants = e_
    Call CopyMember(e_)
    If Alert_Nothing(elem, "fa_Descendants") Then Exit Function
    
    '取得に失敗した場合は一定回数繰り返す
    Dim cnt As Long
    For cnt = 0 To RetryCount
        Set e_.elemArray = elem.FindAllWithOptions(TreeScope_Descendants, c.cnd, direction, elem)
        If e_.elemArray.Length > 0 Then Set e_.elem = e_.elemArray.GetElement(0)
        If Alert_ArrayNothing(e_.elemArray, "fa_Descendants") = False Then Exit For
        Sleep 300
    Next
End Function


'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'TreeWaker関連のラップ
Public Sub tw_Set(twType As EnumTreeWakerType, Optional cnd As IUIAutomationCondition)
    Select Case twType
        Case EnumTreeWakerType.ControlView: Set tw = uia.ControlViewWalker
        Case EnumTreeWakerType.ContentView: Set tw = uia.ContentViewWalker
        Case EnumTreeWakerType.RawView:     Set tw = uia.RawViewWalker
        Case EnumTreeWakerType.ByCondition: Set tw = uia.CreateTreeWalker(cnd)
    End Select
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
    If tw Is Nothing Then Call tw_Set(ControlView)
    Call CopyMember(e_)
    
    Dim strMoveTo As String
    Select Case MoveTo
        Case EnumTreeWakeMoveTo.MoveTo_Parent:      strMoveTo = "tw_Parent"
        Case EnumTreeWakeMoveTo.MoveTo_FirstChild:  strMoveTo = "tw_FirstChild"
        Case EnumTreeWakeMoveTo.MoveTo_LastChild:   strMoveTo = "tw_LastChild"
        Case EnumTreeWakeMoveTo.MoveTo_Next:        strMoveTo = "tw_Next"
        Case EnumTreeWakeMoveTo.MoveTo_Prev:        strMoveTo = "tw_Prev"
    End Select
    
    If Alert_Nothing(elem, strMoveTo & " Root") Then Exit Function
    
    Select Case MoveTo
        Case EnumTreeWakeMoveTo.MoveTo_Parent:      Set e_.elem = tw.GetParentElement(elem)
        Case EnumTreeWakeMoveTo.MoveTo_FirstChild:  Set e_.elem = tw.GetFirstChildElement(elem)
        Case EnumTreeWakeMoveTo.MoveTo_LastChild:   Set e_.elem = tw.GetLastChildElement(elem)
        Case EnumTreeWakeMoveTo.MoveTo_Next:        Set e_.elem = tw.GetNextSiblingElement(elem)
        Case EnumTreeWakeMoveTo.MoveTo_Prev:        Set e_.elem = tw.GetPreviousSiblingElement(elem)
    End Select
    
    Call Alert_Nothing(e_.elem, strMoveTo & " Target")

End Function

'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'エレメントのプロパティのラップ
Public Property Get prName() As String
    If Alert_Nothing(elem, "prName") Then Exit Property
    prName = elem.CurrentName
End Property

Public Property Get prClsName() As String
    If Alert_Nothing(elem, "prClsName") Then Exit Property
    prClsName = elem.CurrentClassName
End Property

Public Property Get prCtrlType() As Long
    If Alert_Nothing(elem, "prCtrlType") Then Exit Property
    prCtrlType = elem.CurrentControlType
End Property

Public Property Get prLocalCtrlType() As String
    If Alert_Nothing(elem, "prLocalCtrlType") Then Exit Property
    prLocalCtrlType = elem.CurrentLocalizedControlType
End Property

Public Property Get prHwnd() As LongPtr
    If Alert_Nothing(elem, "prHwnd") Then Exit Property
    prHwnd = elem.GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId)
End Property

Public Property Get prValue() As String
    If Alert_Nothing(elem, "prValue") Then Exit Property
    prValue = elem.GetCurrentPropertyValue(UIA_ValueValuePropertyId)
End Property

Public Property Get prToggleState() As ToggleState
    If Alert_Nothing(elem, "prToggleState") Then Exit Property
    If Alert_NoPattern(elem, "prToggleState", Toggle) Then Exit Property
    prToggleState = elem.GetCurrentPropertyValue(UIA_ToggleToggleStatePropertyId)
End Property

Public Property Get prExpandCollapseState() As ExpandCollapseState
    If Alert_Nothing(elem, "prExpandCollapseState") Then Exit Property
    If Alert_NoPattern(elem, "prExpandCollapseState", ExpandCollapse) Then Exit Property
    prExpandCollapseState = elem.GetCurrentPropertyValue(UIA_ExpandCollapseExpandCollapseStatePropertyId)
End Property

Public Property Get prScrollState_X() As Double
    If Alert_Nothing(elem, "prScrollState") Then Exit Property
    If Alert_NoPattern(elem, "prScrollState", Scroll) Then Exit Property
    prScrollState_X = elem.GetCurrentPropertyValue(UIA_ScrollHorizontalScrollPercentPropertyId)
End Property

Public Property Get prScrollState_Y() As Double
    If Alert_Nothing(elem, "prScrollState") Then Exit Property
    If Alert_NoPattern(elem, "prScrollState", Scroll) Then Exit Property
    prScrollState_Y = elem.GetCurrentPropertyValue(UIA_ScrollVerticalScrollPercentPropertyId)
End Property

Public Property Get prRect() As tagRECT
    If Alert_Nothing(elem, "prRect") Then Exit Property
    prRect = elem.CurrentBoundingRectangle
End Property

Public Property Get prRectLeft() As Long
    If Alert_Nothing(elem, "prRectLeft") Then Exit Property
    prRectLeft = elem.CurrentBoundingRectangle.Left
End Property

Public Property Get prRectTop() As Long
    If Alert_Nothing(elem, "prRectTop") Then Exit Property
    prRectTop = elem.CurrentBoundingRectangle.Top
End Property

Public Property Get prRectRight() As Long
    If Alert_Nothing(elem, "prRectRight") Then Exit Property
    prRectRight = elem.CurrentBoundingRectangle.Right
End Property

Public Property Get prRectBottom() As Long
    If Alert_Nothing(elem, "prRectBottom") Then Exit Property
    prRectBottom = elem.CurrentBoundingRectangle.bottom
End Property

Public Property Get prRectCenter() As tagPOINT
    If Alert_Nothing(elem, "prRectCenter") Then Exit Property
    prRectCenter.x = (elem.CurrentBoundingRectangle.Left + elem.CurrentBoundingRectangle.Right) / 2
    prRectCenter.y = (elem.CurrentBoundingRectangle.Top + elem.CurrentBoundingRectangle.bottom) / 2
End Property

'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'エレメントのメソッドのラップ
Public Sub SetFocus()
    If Alert_Nothing(elem, "SetFocus") Then Exit Sub
    elem.SetFocus
End Sub
Public Sub ShowContextmenu()
    If Alert_Nothing(elem, "ShowContextMenu") Then Exit Sub
    elem.ShowContextmenu
End Sub

'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'パターン関連のラップ
Public Sub ptInvoke()
    If Alert_Nothing(elem, "ptInvoke") Then Exit Sub
    Dim ptn As IUIAutomationInvokePattern
    If Alert_NoPattern(elem, ptn, Invoke) Then Exit Sub
    ptn.Invoke
End Sub

Public Sub ptSetValue(value As String)
    If Alert_Nothing(elem, "ptSetValue") Then Exit Sub
    Dim ptn As IUIAutomationValuePattern
    If Alert_NoPattern(elem, ptn, Value_) Then Exit Sub
    ptn.setValue value
End Sub

Public Sub ptSelectionItem_Select()
    If Alert_Nothing(elem, "ptSelectionItem_Select") Then Exit Sub
    Dim ptn As IUIAutomationSelectionItemPattern
    If Alert_NoPattern(elem, ptn, SelectionItem) Then Exit Sub
    ptn.Select
End Sub

Public Function ptSelection_GetSelection() As uia_e
    If Alert_Nothing(elem, "ptSelection_GetSelection") Then Exit Function
    Dim ptn As IUIAutomationSelectionPattern
    If Alert_NoPattern(elem, ptn, Selection_) Then Exit Function
    Dim Target As New uia_e
    Set Target.elemArray = ptn.GetCurrentSelection
    If Target.elemArray.Length > 0 Then
        Set Target.elem = Target.elemArray.GetElement(0)
    End If
    Set ptSelection_GetSelection = Target
End Function

Public Sub ptToggle()
    If Alert_Nothing(elem, "ptToggle") Then Exit Sub
    Dim ptn As IUIAutomationTogglePattern
    If Alert_NoPattern(elem, ptn, Toggle) Then Exit Sub
    ptn.Toggle
End Sub

Public Sub ptExpand()
    If Alert_Nothing(elem, "ptExpand") Then Exit Sub
    Dim ptn As IUIAutomationExpandCollapsePattern
    If Alert_NoPattern(elem, ptn, ExpandCollapse) Then Exit Sub
    ptn.Expand
End Sub

Public Sub ptCollapse()
    If Alert_Nothing(elem, "ptExpand") Then Exit Sub
    Dim ptn As IUIAutomationExpandCollapsePattern
    If Alert_NoPattern(elem, ptn, ExpandCollapse) Then Exit Sub
    ptn.Collapse
End Sub

Public Sub ptScroll(Optional XPercent As Double = 99999, Optional YPercent As Double = 99999)
    If Alert_Nothing(elem, "ptScroll") Then Exit Sub
    Dim ptn As IUIAutomationScrollPattern
    If Alert_NoPattern(elem, ptn, Scroll) Then Exit Sub
    Dim scrollX As Long
    Dim scrollY As Long
    If XPercent = 99999 Then
        scrollX = ptn.CurrentHorizontalScrollPercent
    Else
        scrollX = XPercent
    End If
    If YPercent = 99999 Then
        scrollY = ptn.CurrentVerticalScrollPercent
    Else
        scrollY = YPercent
    End If
    ptn.SetScrollPercent scrollX, scrollY
End Sub

Public Sub ptScrollIntoView()
    If Alert_Nothing(elem, "ptScrollIntoView") Then Exit Sub
    Dim ptn As IUIAutomationScrollItemPattern
    If Alert_NoPattern(elem, ptn, ScrollItem) Then Exit Sub
    ptn.ScrollIntoView
End Sub

Public Function ptGridGetItem(row As Long, col As Long) As uia_e
    If Alert_Nothing(elem, "ptGridGetItem") Then Exit Function
    Dim ptn As IUIAutomationGridPattern
    If Alert_NoPattern(elem, ptn, Grid) Then Exit Function
    Dim e_GridItem As New uia_e
    Set e_GridItem.elem = ptn.GetItem(row, col)
    Set ptGridGetItem = e_GridItem
End Function

Public Sub ptWindowClose()
    If Alert_Nothing(elem, "ptWindowClose") Then Exit Sub
    Dim ptn As IUIAutomationWindowPattern
    If Alert_NoPattern(elem, ptn, Window) Then Exit Sub
    ptn.Close
End Sub

Public Function ptGetTextRange() As uia_t
    Dim t_ As New uia_t
    Set ptGetTextRange = t_
    
    If Alert_Nothing(elem, "ptGetTextRange") Then Exit Function
    
    Dim rng As IUIAutomationTextRange
    Select Case True
        Case elem.GetCurrentPropertyValue(UIA_IsTextPattern2AvailablePropertyId)
            Dim TxtPtn2 As IUIAutomationTextPattern2
            Set TxtPtn2 = elem.GetCurrentPattern(UIA_TextPattern2Id)
            Set rng = TxtPtn2.DocumentRange
        Case elem.GetCurrentPropertyValue(UIA_IsTextPatternAvailablePropertyId)
            Dim TxtPtn As IUIAutomationTextPattern
            Set TxtPtn = elem.GetCurrentPattern(UIA_TextPatternId)
            Set rng = TxtPtn.DocumentRange
        Case elem.GetCurrentPropertyValue(UIA_IsTextChildPatternAvailablePropertyId)
            Dim TxtChildPtn As IUIAutomationTextChildPattern
            Set TxtChildPtn = elem.GetCurrentPattern(UIA_TextChildPatternId)
            Set rng = TxtChildPtn.TextRange
        Case elem.GetCurrentPropertyValue(UIA_IsTextEditPatternAvailablePropertyId)
            Dim TxtEditPtn As IUIAutomationTextEditPattern
            Set TxtEditPtn = elem.GetCurrentPattern(UIA_TextEditPatternId)
            Set rng = TxtEditPtn.DocumentRange
        Case Else
            #If DebugMode = 1 Then
                Debug.Print "Any TextPattern Is Not Available! ptGetTextRange"
            #End If
            Exit Function
    End Select

    Set t_.tRng = rng
    rng.ScrollIntoView True
    rng.Select
    
End Function


'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'elemArrayを扱う処理
Public Function Child_SetElemArray(Optional c As uia_c)
    Dim cnd_ As IUIAutomationCondition
    If c Is Nothing Then
        Set cnd_ = uia.CreateTrueCondition
    Else
        If Not c.cnd Is Nothing Then
            Set cnd_ = c.cnd
        End If
    End If
    Set elemArray = elem.FindAll(TreeScope_Children, cnd_)
End Function

Public Function Child_GetItem(c As uia_c) As uia_e
    Dim member() As New uia_e
    Dim buf As uia_e
    ReDim member(elemArray.Length - 1)
    Dim i As Long
    For i = 0 To elemArray.Length - 1
        Set member(i).elem = elemArray.GetElement(i)
        Set buf = member(i).ffDescendants(c, 0)
        If Not Alert_Nothing(buf.elem) Then
            Set Child_GetItem = member(i)
            Exit Function
        End If
    Next
End Function

Public Function Array_GetItemByIndex(Index As Long) As uia_e
    Dim ret_e As New uia_e
    If elemArray Is Nothing Then
        Debug.Print "elemArray Is Nothing! Array_GetItemByIndex(" & Index & ")"
        Exit Function
    End If
    If elemArray.Length - 1 < Index Then
        Debug.Print "elemArray MaxCount Over! Array_GetItemByIndex(" & Index & ")"
        Exit Function
    End If
    
    Set ret_e.elem = elemArray.GetElement(Index)
    Set Array_GetItemByIndex = ret_e
End Function

Public Property Get Array_Length() As Long
    If elemArray Is Nothing Then
        Debug.Print "elemArray Is Nothing! Array_Length"
        Exit Property
    Else
        Array_Length = elemArray.Length
    End If
End Property

Public Property Get ElemArray_Col() As Collection
    If elemArray Is Nothing Then
        Debug.Print "elemArray Is Nothing! ElemArrayCol"
        Exit Property
    Else
        Dim col As New Collection
        Dim i As Long, ChildItem As uia_e
        For i = 0 To elemArray.Length - 1
            Set ChildItem = New uia_e
            CopyMember ChildItem
            Set ChildItem.elem = elemArray.GetElement(i)
            col.Add ChildItem
        Next
        Set ElemArray_Col = col
    End If
End Property

'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'応用的な処理
Function EdgeGetTopWindow(Optional AdditionalCondition As Variant) As uia_e

    Dim e As New uia_e
    Dim c As New uia_c
    Set c = c.ClsName("Chrome_WidgetWin_1").Name1_Sub("edge")

    Select Case typeName(AdditionalCondition)
        Case "String"
            Dim str As String
            str = AdditionalCondition
            Set c = c.Name1_Sub(str)
        Case "Unknown" 'IUIAutomationConditionならコレになる
            Set c.cnd = uia.CreateAndCondition(c.cnd, AdditionalCondition)
        Case "uia_c"
            Set c.cnd = uia.CreateAndCondition(c.cnd, AdditionalCondition.cnd)
    End Select

    Set EdgeGetTopWindow = e.getRoot.ffChildren(c)
    
End Function

Function EdgeGetTopWindows(Optional AdditionalCondition As Variant) As uia_e

    Dim e As New uia_e
    Dim c As New uia_c
    Set c = c.ClsName("Chrome_WidgetWin_1").Name1_Sub("edge")
    
    Set c = CreateAdditionalCondition(c, AdditionalCondition)

    Set EdgeGetTopWindows = e.getRoot.faChildren(c)
    
End Function

Function EdgeGetTabItems(Optional AdditionalCondition As Variant) As uia_e

    'Edgeトップウィンドウが選択されている時専用のコマンド
    If EdgeIsEdgeWindow = False Then Exit Function

    '現在のelemがEdgeのトップウィンドウである前提で処理を進める
    If Alert_Nothing(elem, "EdgeGetTabItems") Then Exit Function

    'Edgeのトップウィンドウからタブ格納領域を取得する
    Dim e_TabContainer As New uia_e
    Dim c_TabContainer As New uia_c
    Set c_TabContainer = c_TabContainer.ClsName("EdgeTabContainerImpl")
    Set e_TabContainer = Me.ffDescendants(c_TabContainer)

    'Edgeのタブ格納領域内にあるタブアイテムを取得する
    Dim e_TabItems As New uia_e
    Dim c_TabItems As New uia_c
    Set c_TabItems = CreateAdditionalCondition(c_TabItems.ClsName("EdgeTab"), AdditionalCondition)
    Set e_TabItems = e_TabContainer.faChildren(c_TabItems)
    
    Set EdgeGetTabItems = e_TabItems

End Function

Function EdgeChangeTab(AdditionalCondition As Variant)
    'Edgeトップウィンドウが選択されている時専用のコマンド
    If EdgeIsEdgeWindow = False Then Exit Function
    
    EdgeGetTabItems(AdditionalCondition).ptSelectionItem_Select
End Function

Function EdgeCreateNewTab() As uia_e
    'Edgeトップウィンドウが選択されている時専用のコマンド
    If EdgeIsEdgeWindow = False Then Exit Function
    
    'タブバー領域内のボタンを押下
    ffDescendants(c.ClsName("EdgeTabStripRegionView")).ffChildren(c.Type_(Button).Name4_Full("新しいタブ")).ptInvoke
    
    Set EdgeCreateNewTab = Me
    
End Function

Function EdgeGetAddresBar() As uia_e
    'Edgeトップウィンドウが選択されている時専用のコマンド
    If EdgeIsEdgeWindow = False Then Exit Function
    
    Set EdgeGetAddresBar = ffDescendants(c.ClsName("OmniboxViewViews").Name4_Full("アドレスと検索バー"))

End Function

Function EdgeSetAddress(value As String, Optional PushEnterKey As Boolean) As uia_e
    'Edgeトップウィンドウが選択されている時専用のコマンド
    EdgeGetAddresBar.ptSetValue value
    
    If PushEnterKey = True Then SendKeys "~"
    
End Function

Private Function EdgeIsEdgeWindow() As Boolean
    'Edgeトップウィンドウが選択されているか判定
    If prClsName = "Chrome_WidgetWin_1" Then
        If prName Like "*Edge" Then
            
            EdgeIsEdgeWindow = True
            
        Else
            Debug.Print "Edgeウィンドウではありません"
        End If
    Else
        Debug.Print "Edgeウィンドウではありません"
    End If
End Function

Function IES_GetIHTMLDocument_FromTopWindow() As IHTMLDocument2
    
    Set IES_GetIHTMLDocument_FromTopWindow = M_LIB_IES.GetIES_From_NomalElement(elem)
    
End Function

Function IES_GetIHTMLDocument_FromIESElement() As IHTMLDocument2
    
    Set IES_GetIHTMLDocument_FromTopWindow = M_LIB_IES.GetIES_From_ieElement(elem)
    
End Function

Private Function CreateAdditionalCondition(c As uia_c, AdditionalCondition As Variant) As uia_c

    Dim cc As New uia_c
    
    Select Case typeName(AdditionalCondition)
        Case "String"
            Dim str As String
            str = AdditionalCondition
            Set cc = c.Name1_Sub(str)
        Case "Unknown" 'IUIAutomationConditionならコレになる
            Set cc.cnd = uia.CreateAndCondition(c.cnd, AdditionalCondition)
        Case "uia_c"
            Set cc.cnd = uia.CreateAndCondition(c.cnd, AdditionalCondition.cnd)
        Case "Error"
            Set cc = c
    End Select
    
    Set CreateAdditionalCondition = cc

End Function

Public Function ptValue_DLFile(Optional DownLoadFolderPath As String, Optional SavedFileFullPath As String, Optional Overwrite As Boolean) As Boolean
    
    Const S_OK As Long = 0
    Dim result As Long

    If DownLoadFolderPath = "" Then
        DownLoadFolderPath = Environ("USERPROFILE") & "\Downloads"
    End If

    If elem.GetCurrentPropertyValue(UIA_IsValuePatternAvailablePropertyId) = False Then
        #If DebugMode = 1 Then
            Debug.Print "No ValuePattern Available!(ptValue_DLFile)"
        #End If
        Exit Function
    End If
    
    Dim fileUrl As String
    fileUrl = elem.GetCurrentPropertyValue(UIA_ValueValuePropertyId)
    If fileUrl = "" Then
        #If DebugMode = 1 Then
            Debug.Print "ValuePattern.Value Is NullString!(ptValue_DLFile)"
        #End If
        Exit Function
    End If
    
    Dim fileNameStartPosition As Long
    fileNameStartPosition = InStrRev(fileUrl, "/")
    If fileNameStartPosition = 0 Then
        #If DebugMode = 1 Then
            Debug.Print """" & fileUrl & """ Is Invalid Url!"
        #End If
        Exit Function
    End If

    Dim fileName As String
    fileName = Mid(fileUrl, fileNameStartPosition + 1)
    
    ' 拡張子とベース名に分割
    Dim BaseName As String, extension As String
    If InStrRev(fileName, ".") > 0 Then
        BaseName = Left(fileName, InStrRev(fileName, ".") - 1)
        extension = Mid(fileName, InStrRev(fileName, "."))
    Else
        BaseName = fileName
        extension = ""
    End If
    
    Dim newName As String
    newName = fileName
    Dim i As Long
    i = 1
    
    ' ファイルが存在する限り (1), (2), ... を付ける
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Overwrite = False Then
        Do While fso.FileExists(fso.BuildPath(DownLoadFolderPath, newName))
            newName = BaseName & " (" & i & ")" & extension
            i = i + 1
        Loop
    End If
    
    Dim DownloadFullPath As String
    DownloadFullPath = fso.BuildPath(DownLoadFolderPath, newName)
    
    result = URLDownloadToFile(0, fileUrl, DownloadFullPath, 0, 0)
    
    ptValue_DLFile = (result = S_OK)
    
    If ptValue_DLFile = False Then
        #If DebugMode = 1 Then
            Debug.Print "File DownLoad Failed!(ptValue_DLFile)"
        #End If
    Else
        SavedFileFullPath = DownloadFullPath
    End If
    
End Function

'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'引き継ぎ用
Private Sub CopyMember(e_ As uia_e)
    If Not tw Is Nothing Then Set e_.tw = tw
    e_.TestMode = Me.TestMode
End Sub

'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'デバッグ用
Private Function Alert_Nothing(elem_ As IUIAutomationElement, Optional FunctionName As String) As Boolean
        
    If TestMode = True Then Exit Function
    
    If elem_ Is Nothing Then
        #If DebugMode = 1 Then
            Debug.Print "elem Is Nothing!(" & FunctionName & ")"
        #End If
        Alert_Nothing = True
    Else
        #If DebugMode2 = 1 Then
            Debug.Print "elem Is Available!(" & FunctionName & ")" & _
            "(" & elem_.CurrentControlType & "/" & GetCtrlTypeName(elem_.CurrentControlType) & ")" & elem_.CurrentName
        #End If
    End If
End Function

Private Function Alert_ArrayNothing(elemArray_ As IUIAutomationElementArray, Optional FunctionName As String) As Boolean
    If elemArray_ Is Nothing Then
        #If DebugMode = 1 Then
            Debug.Print "elemArray Is Nothing!(" & FunctionName & ")"
        #End If
        Alert_ArrayNothing = True
    Else
        #If DebugMode2 = 1 Then
            Debug.Print "elemArray Is Available!(" & FunctionName & ")"
            Dim i As Long
            For i = 0 To elemArray_.Length - 1
                With elemArray_.GetElement(i)
                    
                    Debug.Print i, "(" & .CurrentControlType & "/" & GetCtrlTypeName(.CurrentControlType) & ")" & .CurrentName
                    
                End With
            Next
            
        #End If
    End If
End Function

Private Function Alert_NoPattern(elem_ As IUIAutomationElement, ptn As Variant, PtnId As EnumPattern) As Boolean
    If TestMode = True Then Exit Function
    Set ptn = elem.GetCurrentPattern(PtnId)
    If ptn Is Nothing Then
        Alert_NoPattern = True
        #If DebugMode = 1 Then
            Debug.Print "Pattern Is Nothing! " & uia.GetPatternProgrammaticName(PtnId)
        #End If
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

    Dim res As String
    Select Case elem.CurrentControlType
        Case UIA_AppBarControlTypeId:       res = "AppBar"
        Case UIA_ButtonControlTypeId:       res = "Button"
        Case UIA_CalendarControlTypeId:     res = "Calender"
        Case UIA_CheckBoxControlTypeId:     res = "CheckBox"
        Case UIA_ComboBoxControlTypeId:     res = "ComboBox"
        Case UIA_CustomControlTypeId:       res = "Custom"
        Case UIA_DataGridControlTypeId:     res = "DataGrid"
        Case UIA_DataItemControlTypeId:     res = "DataItem"
        Case UIA_DocumentControlTypeId:     res = "Document"
        Case UIA_EditControlTypeId:         res = "Edit"
        Case UIA_GroupControlTypeId:        res = "Group"
        Case UIA_HeaderControlTypeId:       res = "Header"
        Case UIA_HeaderItemControlTypeId:   res = "HeaderIt"
        Case UIA_HyperlinkControlTypeId:    res = "HyperLnk"
        Case UIA_ImageControlTypeId:        res = "Image"
        Case UIA_ListControlTypeId:         res = "List"
        Case UIA_ListItemControlTypeId:     res = "ListItem"
        Case UIA_MenuBarControlTypeId:      res = "MenuBar"
        Case UIA_MenuControlTypeId:         res = "Menu"
        Case UIA_MenuItemControlTypeId:     res = "MenuItem"
        Case UIA_PaneControlTypeId:         res = "Pane"
        Case UIA_ProgressBarControlTypeId:  res = "Progress"
        Case UIA_RadioButtonControlTypeId:  res = "RadioBtn"
        Case UIA_ScrollBarControlTypeId:    res = "Scroll"
        Case UIA_SemanticZoomControlTypeId: res = "SmtcZoom"
        Case UIA_SeparatorControlTypeId:    res = "Separatr"
        Case UIA_SliderControlTypeId:       res = "Slider"
        Case UIA_SpinnerControlTypeId:      res = "Spinner"
        Case UIA_SplitButtonControlTypeId:  res = "SplitBtn"
        Case UIA_StatusBarControlTypeId:    res = "Status"
        Case UIA_TabControlTypeId:          res = "Tab"
        Case UIA_TabItemControlTypeId:      res = "TabItem"
        Case UIA_TableControlTypeId:        res = "Table"
        Case UIA_TextControlTypeId:         res = "Text"
        Case UIA_ThumbControlTypeId:        res = "Thumb"
        Case UIA_TitleBarControlTypeId:     res = "TitleBar"
        Case UIA_ToolBarControlTypeId:      res = "ToolBar"
        Case UIA_ToolTipControlTypeId:      res = "ToolTip"
        Case UIA_TreeControlTypeId:         res = "Tree"
        Case UIA_TreeItemControlTypeId:     res = "TreeItem"
        Case UIA_WindowControlTypeId:       res = "Window"
        Case Else:                          res = "UnKnown!"
    End Select

    prShortCtrlTypeName = res
    
End Property

'現在のカーソル位置のエレメントを取得するメソッド
Public Function GetFromCursor() As uia_e

    Dim pt As PointAPI
    GetCursorPos pt
    'Debug.Print pt.X & "," & pt.Y
    
    Dim elem As IUIAutomationElement
    Set elem = ElementFromPoint(pt.x, pt.y)

    If Not elem Is Nothing Then
        Set GetFromCursor = New uia_e
        CopyMember GetFromCursor
        Set GetFromCursor.elem = elem
        'Debug.Print "(" & Hex(elem.CurrentControlType) & ")" & elem.CurrentName
    End If

End Function

Public Function GetFromPoint(x As Long, y As Long) As uia_e
    Dim elem As IUIAutomationElement
    Set elem = ElementFromPoint(x, y)

    If Not elem Is Nothing Then
        
        Set GetFromPoint.elem = elem
        'Debug.Print "(" & Hex(elem.CurrentControlType) & ")" & elem.CurrentName
    End If
End Function

'指定された座標のエレメントを取得するメソッド
Private Function ElementFromPoint(x As Long, y As Long) As IUIAutomationElement

    Dim pt As PointAPI
    pt.x = x
    pt.y = y

    Dim Element As IUIAutomationElement
    Dim vParams(pCount) As Variant     'ElementFromPointに渡す各種引数を、バリアント型の配列として準備
    Dim vParamPtr(pCount) As LongPtr   'ElementFromPointに使う各種引数のポインタを、配列として準備
    Dim vParamType(pCount) As Integer  'ElementFromPointに使う各種引数の型を示す値を、Integer型の配列として準備


    '引数本体の格納処理（32bitと64bitで処理が分岐）
    #If Win64 Then
    
        '64bit版のExcelではWOW64を通さず直接関数が呼び出される。
        'stdcall呼び出し規約も無視されるため、引数はCPU内ではスタック領域ではなくレジスタにそのまま放り込まれる。
        'その結果、CUIAutomationのインスタンス領域の不適切なアドレスに値が渡されてメモリ破壊が起こりExcelがクラッシュするリスクがある。
        'これを防ぐため、事前にPOINTAPIの各メンバを、呼び出し先で格納されるtagPOINTと互換性がある単一の変数に手動で格納しておく必要がある。
        Dim llpt As LongLong
        llpt = pt.y * (2 ^ 32) + pt.x 'Shift y left by 4 bits and then add x to put the two parameters into one variable
        
        vParams(0) = llpt
        vParams(1) = VarPtr(Element)
    
    #Else
    
        '32bit環境ではWOW64の中間処理で構造体の各メンバに適切に値が渡されるため、渡し先の仕様を配慮する必要が無い
        vParams(0) = pt.x
        vParams(1) = pt.y
        vParams(2) = VarPtr(Element)
    
    #End If

    '引数の情報（型とアドレス）の格納処理
    Dim pIndex As Long
    For pIndex = 0 To pCount
        vParamPtr(pIndex) = VarPtr(vParams(pIndex))   'ElementFromPointに使う各種引数のポインタ
        vParamType(pIndex) = VarType(vParams(pIndex)) 'ElementFromPointに使う各種引数の型を示す値
    Next

    Dim lRtn As Variant  'DispCallFuncの成否判定を格納する変数
    Dim vRtn As Variant  'ElementFromPointの成否判定を格納する変数
    Dim uRtn As Variant  'IUnknownReleaseの成否判定を格納する変数

    lRtn = DispCallFunc(ObjPtr(uia), pElementFromPoint, CC_STDCALL, vbLong, pCount, vParamType(0), vParamPtr(0), vRtn)
    '　　　　　　　　　　　ByVal　　　　　ByVal　　　    　ByVal　 　ByVal 　ByVal　　　ByRef　　　　ByRef　　　 ByRef
                
                
    '＜DispCallFuncがどのようにしてElementFromPointを呼び出しているのかの解読＞
    
    'CUIAutomationクラスの、
    '仮想関数テーブルの、8番目のポインタが示すアドレスに展開されている処理を実行してください。
    
    '呼び出し規約は「CC_STDCALL」です。※64bit版のExcelでは、この値は無視される

    '実行した処理の戻り値（ElementFromPointの成否判定）はLong型で受け取ります。

    '引数は２つあります。（値渡し）

    '引数の型の種類が書かれたリストはここにあります。（参照渡し）

    '引数が入っている場所が書かれたリストはここにあります。（参照渡し）

    '実行した処理の戻り値（ElementFromPointの成否判定）はここ(vRtn)に格納してください。（参照渡し）
    
                
    If lRtn = S_OK Then
        If vRtn = S_OK Then
            If Not Element Is Nothing Then
                Set ElementFromPoint = Element
                Set Element = Nothing
                'ElementをIUnknown::Releaseで解放しようとすると、
                '呼出元が参照しようとするインターフェースも壊れてExcelがクラッシュするようなので注意'
            End If
        Else
            Debug.Print "ElementFromPoint失敗"
        End If
    Else
        Debug.Print "DispCallFuncメソッド失敗"
    End If
    
End Function

Function GetTopWindow() As uia_e
    
    If Alert_Nothing(Me.elem, "GetTopWindow") Then Exit Function
    
    Dim e As New uia_e
    Set e.elem = Me.elem
    CopyMember e
    Do
        If e.twParent.prName = "デスクトップ 1" Then Exit Do
        Set e = e.twParent
        If Alert_Nothing(e.elem, "GetTopWindow") Then Exit Do
        'Debug.Print e.prName
        DoEvents
    Loop
    Set GetTopWindow = e
End Function

Sub test()

    TestMode = True

    Dim e As New uia_e
    e.TestMode = True

    Call イミディエイトウィドウの末尾にカーソルをセットする
    Debug.Print "取得したい要素にカーソルを合わせてCTRLキーを押してください"
    Debug.Print "終了するにはAltキーを押してください。"
    
    Do
        If (GetAsyncKeyState(vbKeyControl) And -32768) = -32768 Then
            Set e = e.GetFromCursor
            e.GetInfo
            
            Call イミディエイトウィドウの末尾にカーソルをセットする
            Debug.Print "取得したい要素にカーソルを合わせてCTRLキーを押してください"
            Debug.Print "終了するにはALTキーを押してください。"
        Else
            If (GetAsyncKeyState(vbKeyMenu) And -32768) = -32768 Then
                Call イミディエイトウィドウの末尾にカーソルをセットする
                Debug.Print "処理を終了します。"
                Exit Do
            End If
        End If
        DoEvents
        Sleep 1
    Loop
    
    '取得した要素トップウィンドウを取得
    Dim Top As New uia_e
    Set Top = e.GetTopWindow
    
    TestMode = False

End Sub

Private Function ユーザー入力待機() As Boolean
    
    Call イミディエイトウィドウの末尾にカーソルをセットする
    Debug.Print "取得したい要素にカーソルを合わせてCTRLキーを押してください。"
    Debug.Print "終了するにはALTキーを押してください。"
      
    Do
        If (GetAsyncKeyState(vbKeyControl) And -32768) = -32768 Then
            ユーザー入力待機 = True
            Exit Do
        Else
            If (GetAsyncKeyState(vbKeyEscape) And -32768) = -32768 Then
                ユーザー入力待機 = False
                Call イミディエイトウィドウの末尾にカーソルをセットする
                Debug.Print "処理を終了します。"
                Exit Do
            End If
        End If
        DoEvents
        Sleep 1
    Loop
    
End Function

Public Property Get prPtnAv_Invoke() As Boolean
    prPtnAv_Invoke = elem.GetCurrentPropertyValue(UIA_IsInvokePatternAvailablePropertyId)
End Property

Sub GetInfo()
    
    Call イミディエイトウィドウの末尾にカーソルをセットする
    
    Debug.Print "------------------------------"
    Debug.Print "----------Property------------"
    Debug.Print "------------------------------"
    
    Dim l_Property() As Long, s_Property() As String, i_Property As Long
    Call uia.PollForPotentialSupportedProperties(elem, l_Property, s_Property)
    For i_Property = 0 To UBound(l_Property)
        'Debug.Print s_Property(i_Property)
        'IsXXPatternAvailableを除外
        If Not s_Property(i_Property) Like "*Available" Then
            Dim tname As String
            If l_Property(i_Property) <> 0 Then
            'On Error Resume Next
            tname = typeName(elem.GetCurrentPropertyValue(l_Property(i_Property)))
            'Debug.Print tname
            Dim pname As String
            pname = String(40, " ")
            Mid(pname, 1) = s_Property(i_Property)
            
            Select Case tname
                Case "String"
                    Dim strPValue As String
                    strPValue = elem.GetCurrentPropertyValue(l_Property(i_Property))
                    If strPValue <> "" Then
                        Debug.Print pname & strPValue
                    End If

                Case "Long"
                    Dim lngPValue As Long
                    lngPValue = elem.GetCurrentPropertyValue(l_Property(i_Property))
                    If lngPValue <> 0 Then
                        If pname Like "ControlType*" Then
                            Debug.Print pname & lngPValue & "(" & prShortCtrlTypeName & ")"
                        Else
                            Debug.Print pname & lngPValue
                        End If
                    End If

                Case "Boolean"
                    Dim boolPValue As Boolean
                    boolPValue = elem.GetCurrentPropertyValue(l_Property(i_Property))
                    If boolPValue <> False Then
                        Debug.Print pname & boolPValue
                    End If
                    
                Case "Double"
                    Dim dblPValue As Long
                    dblPValue = elem.GetCurrentPropertyValue(l_Property(i_Property))
                    If dblPValue <> 0 Then
                        Debug.Print pname & dblPValue
                    End If
                    
                Case "Double()"
                    Dim dblArrayPValue() As Double
                    dblArrayPValue = elem.GetCurrentPropertyValue(l_Property(i_Property))
                    If UBound(dblArrayPValue) >= 0 Then
                        Dim strDblArrayPValue() As String
                        ReDim strDblArrayPValue(UBound(dblArrayPValue))
                        Dim cnt_D As Long
                        For cnt_D = 0 To UBound(dblArrayPValue)
                            strDblArrayPValue(cnt_D) = dblArrayPValue(cnt_D)
                        Next
                        Debug.Print pname & Join(strDblArrayPValue, ",")
                    End If
                    
                Case "Long()"
                    Dim lngArrayPValue() As Long
                    lngArrayPValue = elem.GetCurrentPropertyValue(l_Property(i_Property))
                    If UBound(lngArrayPValue) >= 0 Then
                        Dim strlngArrayPValue() As String
                        ReDim strlngArrayPValue(UBound(lngArrayPValue))
                        Dim cnt_l As Long
                        For cnt_l = 0 To UBound(lngArrayPValue)
                            strlngArrayPValue(cnt_l) = lngArrayPValue(cnt_l)
                        Next
                        Debug.Print pname & Join(strlngArrayPValue, ",")
                    End If
                
                Case "Nothing"
                Case "Unknown"
                Case "Empty"
            End Select
            Else
                Debug.Print s_Property(i_Property)
            End If
        End If
    Next

    Debug.Print
    Debug.Print "------------------------------"
    Debug.Print "----------Pattern-------------"
    Debug.Print "------------------------------"

    Dim l_Pattern() As Long, s_Pattern() As String, i_Pattern As Long
    Call uia.PollForPotentialSupportedPatterns(elem, l_Pattern, s_Pattern)
    For i_Pattern = 0 To UBound(l_Pattern)
        Debug.Print s_Pattern(i_Pattern)
    Next

End Sub

Private Sub イミディエイトウィドウの末尾にカーソルをセットする()
    
    'VBEのウィンドウを取得
    Dim hVBE As LongPtr
    hVBE = Application.VBE.MainWindow.hWnd
    
    Dim e As New uia_e
    e.TestMode = Me.TestMode
    Set e = e.getHandle(hVBE)

    Dim c As New uia_c
    Dim イミディエイト As New uia_e
    
    Set イミディエイト = e.ffChildren(c.Type_(Window_).Name4_Full("イミディエイト"))
    
    Set イミディエイト = イミディエイト.twLastChild
    
    Dim t As New uia_t
    Set t = イミディエイト.ptGetTextRange
    
    Call t.Move(AsSelfInstance, g_Document, 1)
    
End Sub

Public Function CompareTo(Target As uia_e) As Boolean
    
    Dim 比較結果 As Boolean
    比較結果 = uia.CompareElements(elem, Target.elem)
    
    CompareTo = 比較結果

End Function
