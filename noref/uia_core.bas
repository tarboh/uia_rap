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
Public Const UIA_BoundingRectanglePropertyId As Long = 30001
Public Const UIA_ControlTypePropertyId As Long = 30003
Public Const UIA_LocalizedControlTypePropertyId As Long = 30004
Public Const UIA_NamePropertyId As Long = 30005
Public Const UIA_AutomationIdPropertyId As Long = 30011
Public Const UIA_ClassNamePropertyId As Long = 30012
Public Const UIA_NativeWindowHandlePropertyId As Long = 30020
Public Const UIA_ValueValuePropertyId As Long = 30045
Public Const UIA_ToggleToggleStatePropertyId As Long = 30086
Public Const UIA_ExpandCollapseExpandCollapseStatePropertyId As Long = 30070
Public Const UIA_ScrollHorizontalScrollPercentPropertyId As Long = 30053
Public Const UIA_ScrollVerticalScrollPercentPropertyId As Long = 30055
Public Const UIA_IsInvokePatternAvailablePropertyId As Long = 30031
Public Const UIA_IsValuePatternAvailablePropertyId As Long = 30043
Public Const UIA_IsTextPatternAvailablePropertyId As Long = 30040
Public Const UIA_IsTextPattern2AvailablePropertyId As Long = 30119
Public Const UIA_IsTextChildPatternAvailablePropertyId As Long = 30136
Public Const UIA_IsTextEditPatternAvailablePropertyId As Long = 30149

' --- ControlType Id ---
Public Const UIA_ButtonControlTypeId As Long = 50000
Public Const UIA_CalendarControlTypeId As Long = 50001
Public Const UIA_CheckBoxControlTypeId As Long = 50002
Public Const UIA_ComboBoxControlTypeId As Long = 50003
Public Const UIA_EditControlTypeId As Long = 50004
Public Const UIA_HyperlinkControlTypeId As Long = 50005
Public Const UIA_ImageControlTypeId As Long = 50006
Public Const UIA_ListItemControlTypeId As Long = 50007
Public Const UIA_ListControlTypeId As Long = 50008
Public Const UIA_MenuControlTypeId As Long = 50009
Public Const UIA_MenuBarControlTypeId As Long = 50010
Public Const UIA_MenuItemControlTypeId As Long = 50011
Public Const UIA_ProgressBarControlTypeId As Long = 50012
Public Const UIA_RadioButtonControlTypeId As Long = 50013
Public Const UIA_ScrollBarControlTypeId As Long = 50014
Public Const UIA_SliderControlTypeId As Long = 50015
Public Const UIA_SpinnerControlTypeId As Long = 50016
Public Const UIA_StatusBarControlTypeId As Long = 50017
Public Const UIA_TabControlTypeId As Long = 50018
Public Const UIA_TabItemControlTypeId As Long = 50019
Public Const UIA_TextControlTypeId As Long = 50020
Public Const UIA_ToolBarControlTypeId As Long = 50021
Public Const UIA_ToolTipControlTypeId As Long = 50022
Public Const UIA_TreeControlTypeId As Long = 50023
Public Const UIA_TreeItemControlTypeId As Long = 50024
Public Const UIA_CustomControlTypeId As Long = 50025
Public Const UIA_GroupControlTypeId As Long = 50026
Public Const UIA_ThumbControlTypeId As Long = 50027
Public Const UIA_DataGridControlTypeId As Long = 50028
Public Const UIA_DataItemControlTypeId As Long = 50029
Public Const UIA_DocumentControlTypeId As Long = 50030
Public Const UIA_SplitButtonControlTypeId As Long = 50031
Public Const UIA_WindowControlTypeId As Long = 50032
Public Const UIA_PaneControlTypeId As Long = 50033
Public Const UIA_HeaderControlTypeId As Long = 50034
Public Const UIA_HeaderItemControlTypeId As Long = 50035
Public Const UIA_TableControlTypeId As Long = 50036
Public Const UIA_TitleBarControlTypeId As Long = 50037
Public Const UIA_SeparatorControlTypeId As Long = 50038
Public Const UIA_SemanticZoomControlTypeId As Long = 50039
Public Const UIA_AppBarControlTypeId As Long = 50040

' --- Pattern Id ---
Public Const UIA_InvokePatternId As Long = 10000
Public Const UIA_SelectionPatternId As Long = 10001
Public Const UIA_ValuePatternId As Long = 10002
Public Const UIA_RangeValuePatternId As Long = 10003
Public Const UIA_ScrollPatternId As Long = 10004
Public Const UIA_ExpandCollapsePatternId As Long = 10005
Public Const UIA_GridPatternId As Long = 10006
Public Const UIA_GridItemPatternId As Long = 10007
Public Const UIA_MultipleViewPatternId As Long = 10008
Public Const UIA_WindowPatternId As Long = 10009
Public Const UIA_SelectionItemPatternId As Long = 10010
Public Const UIA_DockPatternId As Long = 10011
Public Const UIA_TablePatternId As Long = 10012
Public Const UIA_TableItemPatternId As Long = 10013
Public Const UIA_TextPatternId As Long = 10014
Public Const UIA_TogglePatternId As Long = 10015
Public Const UIA_TransformPatternId As Long = 10016
Public Const UIA_ScrollItemPatternId As Long = 10017
Public Const UIA_LegacyIAccessiblePatternId As Long = 10018
Public Const UIA_ItemContainerPatternId As Long = 10019
Public Const UIA_VirtualizedItemPatternId As Long = 10020
Public Const UIA_SynchronizedInputPatternId As Long = 10021
Public Const UIA_ObjectModelPatternId As Long = 10022
Public Const UIA_AnnotationPatternId As Long = 10023
Public Const UIA_TextPattern2Id As Long = 10024
Public Const UIA_StylesPatternId As Long = 10025
Public Const UIA_SpreadsheetPatternId As Long = 10026
Public Const UIA_SpreadsheetItemPatternId As Long = 10027
Public Const UIA_TransformPattern2Id As Long = 10028
Public Const UIA_TextChildPatternId As Long = 10029
Public Const UIA_DragPatternId As Long = 10030
Public Const UIA_DropTargetPatternId As Long = 10031
Public Const UIA_TextEditPatternId As Long = 10032
Public Const UIA_CustomNavigationPatternId As Long = 10033
Public Const UIA_SelectionPattern2Id As Long = 10034

' --- PropertyConditionFlags ---
Public Const PropertyConditionFlags_None As Long = 0
Public Const PropertyConditionFlags_IgnoreCase As Long = 1
Public Const PropertyConditionFlags_MatchSubstring As Long = 2

' --- TreeScope ---
Public Const TreeScope_None As Long = 0
Public Const TreeScope_Element As Long = 1
Public Const TreeScope_Children As Long = 2
Public Const TreeScope_Descendants As Long = 4
Public Const TreeScope_Parent As Long = 8
Public Const TreeScope_Ancestors As Long = 16
Public Const TreeScope_Subtree As Long = 7

' --- TreeTraversalOptions ---
Public Const TreeTraversalOptions_Default As Long = 0
Public Const TreeTraversalOptions_PostOrder As Long = 1
Public Const TreeTraversalOptions_LastToFirstOrder As Long = 2

' --- TextUnit ---
Public Const TextUnit_Character As Long = 0
Public Const TextUnit_Format As Long = 1
Public Const TextUnit_Word As Long = 2
Public Const TextUnit_Line As Long = 3
Public Const TextUnit_Paragraph As Long = 4
Public Const TextUnit_Page As Long = 5
Public Const TextUnit_Document As Long = 6

' --- TextPatternRangeEndpoint ---
Public Const TextPatternRangeEndpoint_Start As Long = 0
Public Const TextPatternRangeEndpoint_End As Long = 1

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
    GetPropertyValue = v
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
