Attribute VB_Name = "uia_core"
'==============================================================================
' uia_core  ―  参照設定不要版の下回り（プラミング）
'
' uia_e / uia_c / uia_t が使う共通基盤をここに集約する。ユーザーが直接触ることは
' 想定していない（インスタンス生成もここに隠蔽する）。
'
'  * CoCreateInstance + DispCallFunc で UI Automation を叩く
'  * 共有の IUIAutomation ポインタ（UIA()）を遅延生成して保持する
'  * BSTR / VARIANT / SAFEARRAY のマーシャリング
'  * uia_c / uia_e が参照する UIA_* 定数
'
' 参照設定は一切不要。VBA7 (Office 2010) 以降、32bit / 64bit 両対応。
'==============================================================================
Option Explicit

'--- CLSID / IID -------------------------------------------------------------
Private Const CLSID_CUIAutomation As String = "{ff48dba4-60ef-4201-aa87-54103eef594e}"
Private Const IID_IUIAutomation As String = "{30cbe57d-d9d0-452a-ab13-7ac5ac4825ee}"
Public Const IID_IUIAutomationElement3 As String = "{8471df34-aee0-4a01-a7de-7db9af12c296}"
Public Const IID_IUIAutomationElement9 As String = "{39325fac-039d-440e-a3a3-5eb81a5cecc3}"

'--- ポインタ幅 / VARENUM -----------------------------------------------------
#If Win64 Then
    Public Const PTR_SIZE As Long = 8
#Else
    Public Const PTR_SIZE As Long = 4
#End If

Public Const VT_I2 As Integer = 2
Public Const VT_I4 As Integer = 3
Public Const VT_R4 As Integer = 4
Public Const VT_R8 As Integer = 5
Public Const VT_BSTR As Integer = 8
Public Const VT_BOOL As Integer = 11
Public Const VT_VARIANT As Integer = 12
Public Const VT_UI4 As Integer = 19
Public Const VT_I8 As Integer = 20
Public Const VT_UNKNOWN As Integer = 13

Private Const CC_STDCALL As Long = 4
Private Const CLSCTX_INPROC_SERVER As Long = 1
Public Const S_OK As Long = 0
Public Const UIA_ERR_BASE As Long = vbObjectError + &H5100&

'--- 構造体 ------------------------------------------------------------------
Public Type UIA_GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

' UIAutomation クライアント API の RECT は Win32 の RECT (LONG x4)
Public Type tagRECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type tagPOINT
    X As Long
    Y As Long
End Type

'==============================================================================
' UIA_* 定数（uia_c / uia_e が参照）
'==============================================================================

' --- Property Id ---
Public Enum UIA_PropertyIDs
    UIA_BoundingRectanglePropertyId = 30001
    UIA_ControlTypePropertyId = 30003
    UIA_LocalizedControlTypePropertyId = 30004
    UIA_NamePropertyId = 30005
    UIA_AutomationIdPropertyId = 30011
    UIA_ClassNamePropertyId = 30012
    UIA_NativeWindowHandlePropertyId = 30020
    UIA_ValueValuePropertyId = 30045
    UIA_ToggleToggleStatePropertyId = 30086
    UIA_ExpandCollapseExpandCollapseStatePropertyId = 30070
    UIA_ScrollHorizontalScrollPercentPropertyId = 30053
    UIA_ScrollVerticalScrollPercentPropertyId = 30055
    UIA_IsInvokePatternAvailablePropertyId = 30031
    UIA_IsValuePatternAvailablePropertyId = 30043
    UIA_IsTextPatternAvailablePropertyId = 30040
    UIA_IsTextPattern2AvailablePropertyId = 30119
    UIA_IsTextChildPatternAvailablePropertyId = 30136
    UIA_IsTextEditPatternAvailablePropertyId = 30149
End Enum

' --- ControlType Id ---
Public Enum UIA_ControlTypeIDs
    UIA_ButtonControlTypeId = 50000
    UIA_CalendarControlTypeId = 50001
    UIA_CheckBoxControlTypeId = 50002
    UIA_ComboBoxControlTypeId = 50003
    UIA_EditControlTypeId = 50004
    UIA_HyperlinkControlTypeId = 50005
    UIA_ImageControlTypeId = 50006
    UIA_ListItemControlTypeId = 50007
    UIA_ListControlTypeId = 50008
    UIA_MenuControlTypeId = 50009
    UIA_MenuBarControlTypeId = 50010
    UIA_MenuItemControlTypeId = 50011
    UIA_ProgressBarControlTypeId = 50012
    UIA_RadioButtonControlTypeId = 50013
    UIA_ScrollBarControlTypeId = 50014
    UIA_SliderControlTypeId = 50015
    UIA_SpinnerControlTypeId = 50016
    UIA_StatusBarControlTypeId = 50017
    UIA_TabControlTypeId = 50018
    UIA_TabItemControlTypeId = 50019
    UIA_TextControlTypeId = 50020
    UIA_ToolBarControlTypeId = 50021
    UIA_ToolTipControlTypeId = 50022
    UIA_TreeControlTypeId = 50023
    UIA_TreeItemControlTypeId = 50024
    UIA_CustomControlTypeId = 50025
    UIA_GroupControlTypeId = 50026
    UIA_ThumbControlTypeId = 50027
    UIA_DataGridControlTypeId = 50028
    UIA_DataItemControlTypeId = 50029
    UIA_DocumentControlTypeId = 50030
    UIA_SplitButtonControlTypeId = 50031
    UIA_WindowControlTypeId = 50032
    UIA_PaneControlTypeId = 50033
    UIA_HeaderControlTypeId = 50034
    UIA_HeaderItemControlTypeId = 50035
    UIA_TableControlTypeId = 50036
    UIA_TitleBarControlTypeId = 50037
    UIA_SeparatorControlTypeId = 50038
    UIA_SemanticZoomControlTypeId = 50039
    UIA_AppBarControlTypeId = 50040
End Enum

' --- Pattern Id ---
Public Enum UIA_PatternIDs
    UIA_InvokePatternId = 10000
    UIA_SelectionPatternId = 10001
    UIA_ValuePatternId = 10002
    UIA_RangeValuePatternId = 10003
    UIA_ScrollPatternId = 10004
    UIA_ExpandCollapsePatternId = 10005
    UIA_GridPatternId = 10006
    UIA_GridItemPatternId = 10007
    UIA_MultipleViewPatternId = 10008
    UIA_WindowPatternId = 10009
    UIA_SelectionItemPatternId = 10010
    UIA_DockPatternId = 10011
    UIA_TablePatternId = 10012
    UIA_TableItemPatternId = 10013
    UIA_TextPatternId = 10014
    UIA_TogglePatternId = 10015
    UIA_TransformPatternId = 10016
    UIA_ScrollItemPatternId = 10017
    UIA_LegacyIAccessiblePatternId = 10018
    UIA_ItemContainerPatternId = 10019
    UIA_VirtualizedItemPatternId = 10020
    UIA_SynchronizedInputPatternId = 10021
    UIA_ObjectModelPatternId = 10022
    UIA_AnnotationPatternId = 10023
    UIA_TextPattern2Id = 10024
    UIA_StylesPatternId = 10025
    UIA_SpreadsheetPatternId = 10026
    UIA_SpreadsheetItemPatternId = 10027
    UIA_TransformPattern2Id = 10028
    UIA_TextChildPatternId = 10029
    UIA_DragPatternId = 10030
    UIA_DropTargetPatternId = 10031
    UIA_TextEditPatternId = 10032
    UIA_CustomNavigationPatternId = 10033
    UIA_SelectionPattern2Id = 10034
End Enum

' --- PropertyConditionFlags ---
Public Enum PropertyConditionFlags
    PropertyConditionFlags_None = 0
    PropertyConditionFlags_IgnoreCase = 1
    PropertyConditionFlags_MatchSubstring = 2
End Enum

' --- TreeScope ---
Public Enum TreeScope
    TreeScope_None = 0
    TreeScope_Element = 1
    TreeScope_Children = 2
    TreeScope_Descendants = 4
    TreeScope_Parent = 8
    TreeScope_Ancestors = 16
    TreeScope_Subtree = 7
End Enum

' --- TreeTraversalOptions ---
Public Enum TreeTraversalOptions
    TreeTraversalOptions_Default = 0
    TreeTraversalOptions_PostOrder = 1
    TreeTraversalOptions_LastToFirstOrder = 2
End Enum

' --- TextUnit ---
Public Enum TextUnit
    TextUnit_Character = 0
    TextUnit_Format = 1
    TextUnit_Word = 2
    TextUnit_Line = 3
    TextUnit_Paragraph = 4
    TextUnit_Page = 5
    TextUnit_Document = 6
End Enum

' --- TextPatternRangeEndpoint ---
Public Enum TextPatternRangeEndpoint
    TextPatternRangeEndpoint_Start = 0
    TextPatternRangeEndpoint_End = 1
End Enum

'==============================================================================
' API 宣言
'==============================================================================
#If VBA7 Then
Private Declare PtrSafe Function CoCreateInstance Lib "ole32" ( _
    ByRef rclsid As UIA_GUID, ByVal pUnkOuter As LongPtr, ByVal dwClsContext As Long, _
    ByRef riid As UIA_GUID, ByRef ppv As LongPtr) As Long
Private Declare PtrSafe Function CLSIDFromString Lib "ole32" ( _
    ByVal lpsz As LongPtr, ByRef pclsid As UIA_GUID) As Long
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" ( _
    ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As Long, _
    ByVal vtReturn As Integer, ByVal cActuals As Long, _
    ByRef prgvt As Integer, ByRef prgpvarg As LongPtr, ByRef pvargResult As Variant) As Long
Public Declare PtrSafe Function SysStringLen Lib "oleaut32" (ByVal bstr As LongPtr) As Long
Public Declare PtrSafe Sub SysFreeString Lib "oleaut32" (ByVal bstr As LongPtr)
Public Declare PtrSafe Function SafeArrayGetLBound Lib "oleaut32" ( _
    ByVal psa As LongPtr, ByVal nDim As Long, ByRef plLbound As Long) As Long
Public Declare PtrSafe Function SafeArrayGetUBound Lib "oleaut32" ( _
    ByVal psa As LongPtr, ByVal nDim As Long, ByRef plUbound As Long) As Long
Public Declare PtrSafe Function SafeArrayGetElement Lib "oleaut32" ( _
    ByVal psa As LongPtr, ByRef rgIndices As Long, ByRef pv As Any) As Long
Public Declare PtrSafe Function SafeArrayDestroy Lib "oleaut32" (ByVal psa As LongPtr) As Long
Private Declare PtrSafe Sub Sleep_API Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
    ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As LongPtr) As Long
Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare PtrSafe Function GetCursorPos Lib "user32" ( _
    ByRef lpPoint As tagPOINT) As Long
Private Declare PtrSafe Function FormatMessageW Lib "kernel32" ( _
    ByVal dwFlags As Long, ByVal lpSource As LongPtr, ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, ByVal lpBuffer As LongPtr, ByVal nSize As Long, _
    ByVal Arguments As LongPtr) As Long
#Else
' VBA7 未満 (Office 2007 以前) は LongPtr を持たないため未対応。
#End If

'==============================================================================
' 共有 IUIAutomation
'==============================================================================
Private g_uia As LongPtr

' 共有の IUIAutomation ポインタ。初回アクセス時に CoCreateInstance で生成する。
Public Function UIA() As LongPtr
    If g_uia = 0 Then
        g_uia = ComCreateInstance(CLSID_CUIAutomation, IID_IUIAutomation)
    End If
    UIA = g_uia
End Function

' 明示的に解放したいとき用（通常はプロセス終了に任せてよい）。
Public Sub UIA_Shutdown()
    If g_uia <> 0 Then
        ComRelease g_uia
        g_uia = 0
    End If
End Sub

'==============================================================================
' vtable 呼び出しエンジン
'==============================================================================

' UiaInvoke(pThis, vtblIndex, argSpec, arg1, ...) As Long ' 戻り値は HRESULT
'   argSpec: L=Long U=ULong I=Int 8=LongLong(x64) P=ポインタ幅
'            D=Double E=Single S=BSTR B=Boolean V=Variant
Public Function UiaInvoke(ByVal pThis As LongPtr, ByVal vtblIndex As Long, _
                          ByVal argSpec As String, ParamArray args() As Variant) As Long
    Dim n As Long, i As Long
    Dim vArgs() As Variant, vTypes() As Integer, pArgs() As LongPtr
    Dim vRet As Variant, hrCall As Long

    If pThis = 0 Then
        UiaInvoke = &H80004003             ' E_POINTER
        Exit Function
    End If

    n = Len(argSpec)
    ReDim vArgs(0 To IIf(n = 0, 0, n - 1))
    ReDim vTypes(0 To IIf(n = 0, 0, n - 1))
    ReDim pArgs(0 To IIf(n = 0, 0, n - 1))

    For i = 0 To n - 1
        vTypes(i) = CoerceArg(Mid$(argSpec, i + 1, 1), args(LBound(args) + i), vArgs(i))
        pArgs(i) = VarPtr(vArgs(i))
    Next i

    hrCall = DispCallFunc(pThis, vtblIndex * PTR_SIZE, CC_STDCALL, _
                          VT_I4, n, vTypes(0), pArgs(0), vRet)
    If hrCall <> S_OK Then
        UiaInvoke = hrCall
    Else
        UiaInvoke = CLng(vRet)
    End If
End Function

Private Function CoerceArg(ByVal code As String, ByRef src As Variant, ByRef dst As Variant) As Integer
    Select Case code
        Case "L": dst = CLng(src):  CoerceArg = VT_I4
        Case "U": dst = CLng(src):  CoerceArg = VT_UI4
        Case "I": dst = CInt(src):  CoerceArg = VT_I2
        Case "D": dst = CDbl(src):  CoerceArg = VT_R8
        Case "E": dst = CSng(src):  CoerceArg = VT_R4
        Case "S": dst = CStr(src):  CoerceArg = VT_BSTR
        Case "B": dst = CBool(src): CoerceArg = VT_BOOL
        Case "V"
            If IsObject(src) Then Set dst = src Else dst = src
            CoerceArg = VT_VARIANT
        Case "P", "8"
#If Win64 Then
            dst = CLngLng(src): CoerceArg = VT_I8
#Else
            If code = "8" Then Err.Raise UIA_ERR_BASE, "uia_core", "argSpec '8' は 32bit では不可"
            dst = CLng(src):    CoerceArg = VT_I4
#End If
        Case Else
            Err.Raise UIA_ERR_BASE, "uia_core", "未知の argSpec: '" & code & "'"
    End Select
End Function


'==============================================================================
' IUnknown / 生成 / GUID
'==============================================================================
Public Function ComQI(ByVal pUnk As LongPtr, ByVal iidString As String) As LongPtr
    Dim g As UIA_GUID, p As LongPtr
    If pUnk = 0 Then Exit Function
    g = GuidFromString(iidString)
    If UiaInvoke(pUnk, 0, "PP", VarPtr(g), VarPtr(p)) <> S_OK Then p = 0
    ComQI = p
End Function

Public Function ComAddRef(ByVal pUnk As LongPtr) As Long
    If pUnk <> 0 Then ComAddRef = UiaInvoke(pUnk, 1, "")
End Function

Public Function ComRelease(ByVal pUnk As LongPtr) As Long
    If pUnk <> 0 Then ComRelease = UiaInvoke(pUnk, 2, "")
End Function

Public Function ComCreateInstance(ByVal clsidString As String, ByVal iidString As String) As LongPtr
    Dim c As UIA_GUID, i As UIA_GUID, p As LongPtr, hr As Long
    c = GuidFromString(clsidString)
    i = GuidFromString(iidString)
    hr = CoCreateInstance(c, 0, CLSCTX_INPROC_SERVER, i, p)
    If hr <> S_OK Then UiaCheck hr, "CoCreateInstance(" & clsidString & ")"
    ComCreateInstance = p
End Function

Public Function GuidFromString(ByVal s As String) As UIA_GUID
    Dim g As UIA_GUID
    If Left$(s, 1) <> "{" Then s = "{" & s & "}"
    If CLSIDFromString(StrPtr(s), g) <> S_OK Then
        Err.Raise UIA_ERR_BASE, "uia_core.GuidFromString", "GUID 文字列が不正: " & s
    End If
    GuidFromString = g
End Function

'==============================================================================
' 条件生成 (IUIAutomation vtable) ― uia_c から使う
'   戻り値は AddRef 済みの生ポインタ。呼び出し側 (uia_c) が所有・解放する。
'==============================================================================
Public Function CreatePropertyCondition(ByVal propId As Long, ByVal value As Variant) As LongPtr
    Dim p As LongPtr
    UiaCheck UiaInvoke(UIA(), 23, "LVP", propId, value, VarPtr(p)), "CreatePropertyCondition"
    CreatePropertyCondition = p
End Function

Public Function CreatePropertyConditionEx(ByVal propId As Long, ByVal value As Variant, _
                                          ByVal flags As Long) As LongPtr
    Dim p As LongPtr
    UiaCheck UiaInvoke(UIA(), 24, "LVLP", propId, value, flags, VarPtr(p)), _
             "CreatePropertyConditionEx"
    CreatePropertyConditionEx = p
End Function

Public Function CreateAndCondition(ByVal c1 As LongPtr, ByVal c2 As LongPtr) As LongPtr
    Dim p As LongPtr
    UiaCheck UiaInvoke(UIA(), 25, "PPP", c1, c2, VarPtr(p)), "CreateAndCondition"
    CreateAndCondition = p
End Function

Public Function CreateOrCondition(ByVal c1 As LongPtr, ByVal c2 As LongPtr) As LongPtr
    Dim p As LongPtr
    UiaCheck UiaInvoke(UIA(), 28, "PPP", c1, c2, VarPtr(p)), "CreateOrCondition"
    CreateOrCondition = p
End Function

Public Function CreateNotCondition(ByVal c1 As LongPtr) As LongPtr
    Dim p As LongPtr
    UiaCheck UiaInvoke(UIA(), 31, "PP", c1, VarPtr(p)), "CreateNotCondition"
    CreateNotCondition = p
End Function

Public Function CreateTrueCondition() As LongPtr
    Dim p As LongPtr
    UiaCheck UiaInvoke(UIA(), 21, "P", VarPtr(p)), "CreateTrueCondition"
    CreateTrueCondition = p
End Function

Public Function CreateFalseCondition() As LongPtr
    Dim p As LongPtr
    UiaCheck UiaInvoke(UIA(), 22, "P", VarPtr(p)), "CreateFalseCondition"
    CreateFalseCondition = p
End Function

'==============================================================================
' 要素 / パターン / 座標 ヘルパ ― uia_e から使う
'==============================================================================

' 要素のプロパティ値 (VARIANT) を取得する。GetCurrentPropertyValue = vtable[10]。
Public Function GetPropertyValue(ByVal pElem As LongPtr, ByVal propId As Long) As Variant
    Dim v As Variant
    uia_core.UiaCheck UiaInvoke(pElem, 10, "LP", propId, VarPtr(v)), _
                      "IUIAutomationElement.GetCurrentPropertyValue"
    ' 一部のプロパティ (LabeledBy / ControllerFor 等) は要素参照 (VT_UNKNOWN) を返す。
    ' 参照設定が無いと型情報が無く、VBA が扱えない Variant になって代入時に
    ' 「型が一致しません」になる。IsObject では捕まらないケースがあるため、
    ' 代入自体をエラートラップして確実に弾き、そういう値は Empty で返す。
    ' (ローカル v はスコープ終了時に VBA が解放するので参照リークもしない)
    On Error Resume Next
    GetPropertyValue = v
    If Err.Number <> 0 Then
        Err.Clear
        GetPropertyValue = Empty
    End If
    On Error GoTo 0
End Function

' 要素のパターンを取得する。GetCurrentPattern = vtable[16]。未サポートなら 0。
' 戻り値は AddRef 済み。呼び出し側で ComRelease すること。
Public Function GetPattern(ByVal pElem As LongPtr, ByVal patternId As Long) As LongPtr
    Dim p As LongPtr
    uia_core.UiaCheck UiaInvoke(pElem, 16, "LP", patternId, VarPtr(p)), _
                      "IUIAutomationElement.GetCurrentPattern"
    GetPattern = p
End Function

' 要素の矩形。get_CurrentBoundingRectangle = vtable[43]。RECT(LONG x4) を受ける。
Public Function GetBoundingRect(ByVal pElem As LongPtr) As tagRECT
    Dim b(0 To 3) As Long
    uia_core.UiaCheck UiaInvoke(pElem, 43, "P", VarPtr(b(0))), _
                      "IUIAutomationElement.get_CurrentBoundingRectangle"
    GetBoundingRect.Left = b(0)
    GetBoundingRect.Top = b(1)
    GetBoundingRect.Right = b(2)
    GetBoundingRect.Bottom = b(3)
End Function

' [out] Long を1個返すメソッドをラップする。
Public Function InvokeLong(ByVal pThis As LongPtr, ByVal vtblIndex As Long, _
                           ByVal context As String) As Long
    Dim v As Long
    uia_core.UiaCheck UiaInvoke(pThis, vtblIndex, "P", VarPtr(v)), context
    InvokeLong = v
End Function

' [out] インターフェースポインタを1個返すメソッドをラップする。戻り値は AddRef 済み。
Public Function InvokeElem(ByVal pThis As LongPtr, ByVal vtblIndex As Long, _
                           ByVal context As String) As LongPtr
    Dim p As LongPtr
    uia_core.UiaCheck UiaInvoke(pThis, vtblIndex, "P", VarPtr(p)), context
    InvokeElem = p
End Function

' 座標から要素を取得する。ElementFromPoint = IUIAutomation vtable[7]。
' POINT の値渡しは 32/64bit で ABI が違う (CopyMemory は使わず算術で組む)。
Public Function ElementFromPoint(ByVal x As Long, ByVal y As Long) As LongPtr
    Dim p As LongPtr
#If Win64 Then
    Dim packed As LongLong
    packed = (CLngLng(y) * &H100000000^) Or (CLngLng(x) And &HFFFFFFFF^)
    uia_core.UiaCheck UiaInvoke(UIA(), 7, "8P", packed, VarPtr(p)), "ElementFromPoint"
#Else
    uia_core.UiaCheck UiaInvoke(UIA(), 7, "LLP", x, y, VarPtr(p)), "ElementFromPoint"
#End If
    ElementFromPoint = p
End Function

' マウスカーソルの現在位置。
Public Function CursorPos() As tagPOINT
    GetCursorPos CursorPos
End Function

' 指定ミリ秒スリープする。
Public Sub Sleep_(ByVal dwMilliseconds As Long)
    Sleep_API dwMilliseconds
End Sub

'==============================================================================
' 文字列 / 配列
'==============================================================================

' [out] BSTR* をラップ呼び出しする。戻り値の String に COM 側が直接書き込むので
' CopyMemory は不要（VBA が所有・解放する）。idx は BSTR を1個返すメソッドの vtable。
Public Function InvokeBstr(ByVal pThis As LongPtr, ByVal vtblIndex As Long, _
                           ByVal context As String) As String
    Dim hr As Long
    hr = UiaInvoke(pThis, vtblIndex, "P", VarPtr(InvokeBstr))
    UiaCheck hr, context
End Function

' [out] SAFEARRAY*(VT_I4) を Long 配列にする（RuntimeId 等）。
Public Function SafeArrayToLongs(ByVal psa As LongPtr, _
                                 Optional ByVal destroyIt As Boolean = True) As Long()
    Dim lb As Long, ub As Long, i As Long, v As Long, out() As Long
    If psa = 0 Then SafeArrayToLongs = out: Exit Function
    SafeArrayGetLBound psa, 1, lb
    SafeArrayGetUBound psa, 1, ub
    If ub >= lb Then
        ReDim out(0 To ub - lb)
        For i = lb To ub
            SafeArrayGetElement psa, i, v
            out(i - lb) = v
        Next i
    End If
    If destroyIt Then SafeArrayDestroy psa
    SafeArrayToLongs = out
End Function

' [out] SAFEARRAY*(VT_BSTR) を String 配列にする。
' SafeArrayGetElement は VT_BSTR 要素をコピーして String のスロットへ書くので
' CopyMemory は不要 (VBA が所有・解放する)。
Public Function SafeArrayToStrings(ByVal psa As LongPtr, _
                                   Optional ByVal destroyIt As Boolean = True) As String()
    Dim lb As Long, ub As Long, i As Long, out() As String, s As String
    If psa = 0 Then SafeArrayToStrings = out: Exit Function
    SafeArrayGetLBound psa, 1, lb
    SafeArrayGetUBound psa, 1, ub
    If ub >= lb Then
        ReDim out(0 To ub - lb)
        For i = lb To ub
            s = vbNullString                    ' 空にしておく (上書きで旧BSTRを漏らさない)
            SafeArrayGetElement psa, i, s       ' VT_BSTR 要素を s にコピー
            out(i - lb) = s
        Next
    End If
    If destroyIt Then SafeArrayDestroy psa
    SafeArrayToStrings = out
End Function

' 要素がサポートしうるプロパティ ID と名前を取得する (vtable[52])。
Public Sub PollSupportedProperties(ByVal pElem As LongPtr, _
                                   ByRef ids() As Long, ByRef names() As String)
    Dim saIds As LongPtr, saNames As LongPtr
    UiaCheck UiaInvoke(UIA(), 52, "PPP", pElem, VarPtr(saIds), VarPtr(saNames)), _
             "PollForPotentialSupportedProperties"
    ids = SafeArrayToLongs(saIds)
    names = SafeArrayToStrings(saNames)
End Sub

' 要素がサポートしうるパターン ID と名前を取得する (vtable[51])。
Public Sub PollSupportedPatterns(ByVal pElem As LongPtr, _
                                 ByRef ids() As Long, ByRef names() As String)
    Dim saIds As LongPtr, saNames As LongPtr
    UiaCheck UiaInvoke(UIA(), 51, "PPP", pElem, VarPtr(saIds), VarPtr(saNames)), _
             "PollForPotentialSupportedPatterns"
    ids = SafeArrayToLongs(saIds)
    names = SafeArrayToStrings(saNames)
End Sub

' パターン ID の内部名 (vtable[50])。
Public Function GetPatternProgrammaticName(ByVal patternId As Long) As String
    UiaCheck UiaInvoke(UIA(), 50, "LP", patternId, VarPtr(GetPatternProgrammaticName)), _
             "GetPatternProgrammaticName"
End Function

' プロパティ ID の内部名 (vtable[49])。
Public Function GetPropertyProgrammaticName(ByVal propertyId As Long) As String
    UiaCheck UiaInvoke(UIA(), 49, "LP", propertyId, VarPtr(GetPropertyProgrammaticName)), _
             "GetPropertyProgrammaticName"
End Function

'==============================================================================
' 自己テスト（uia_e 移植前の下回り確認用）
'==============================================================================

' 参照設定なしで IUIAutomation が生成でき、条件が組めるかを確認する。
' noref/ の uia_core.bas と uia_c.cls をインポートしてから実行すること。
Public Sub UIA_NoRef_SelfTest()
    ' 1) CoCreateInstance + DispCallFunc が参照なしで通るか
    If UIA() = 0 Then
        Err.Raise UIA_ERR_BASE, "selftest", "IUIAutomation を生成できませんでした。"
    End If
    Debug.Print "UIA() OK  ptr=" & UIA()

    ' 2) 条件ビルダ (VARIANT 引数 + 参照カウント) が動くか
    Dim builder As New uia_c
    Dim cond As uia_c
    Set cond = builder.Type_(UIA_ButtonControlTypeId).Name4_Full("OK")
    If cond.cnd = 0 Then
        Err.Raise UIA_ERR_BASE, "selftest", "条件を生成できませんでした。"
    End If
    Debug.Print "condition OK  ptr=" & cond.cnd

    Debug.Print "uia_rap (no-ref) 下回り 自己テスト: OK"
End Sub

'==============================================================================
' エラー
'==============================================================================
Public Sub UiaCheck(ByVal hr As Long, ByVal context As String)
    If hr >= 0 Then Exit Sub
    Err.Raise UIA_ERR_BASE + (hr And &HFFF&), "uia (noref)", _
              context & " が失敗しました。HRESULT=0x" & Hex$(hr) & _
              IIf(Len(HResultText(hr)) > 0, " (" & HResultText(hr) & ")", "")
End Sub

Public Function Succeeded(ByVal hr As Long) As Boolean
    Succeeded = (hr >= 0)
End Function

Private Function HResultText(ByVal hr As Long) As String
    Dim buf As String, n As Long
    buf = String$(512, vbNullChar)
    n = FormatMessageW(&H1000& Or &H200&, 0, hr, 0, StrPtr(buf), 512, 0)
    If n > 0 Then HResultText = Trim$(Replace$(Replace$(Left$(buf, n), vbCr, ""), vbLf, ""))
End Function
