Attribute VB_Name = "M_LIB_IES"
'==============================================================================
' M_LIB_IES (参照設定不要版)  ―  IE / Edge の IE モードから HTML DOM を取得する
'
' Edge の「IE モード」タブや IE の "Internet Explorer_Server" クラスの要素から、
' IHTMLDocument2 を取り出す。参照設定なしで動くよう、要素は生ポインタ (LongPtr)
' で受け取り、戻り値は遅延バインドの Object (実体は IHTMLDocument2) で返す。
'
' 使い方 (uia_e 経由):
'   Dim doc As Object
'   Set doc = edgeTab.IES_GetIHTMLDocument_FromTopWindow()
'   Debug.Print doc.title
'   doc.getElementById("foo").Click
'
' Object 遅延バインドなので IntelliSense は効かないが、IE モードでは実用になる。
' Microsoft HTML Object Library の参照は不要。
'==============================================================================
Option Explicit

Private Declare PtrSafe Function RegisterWindowMessage Lib "user32" _
    Alias "RegisterWindowMessageA" (ByVal lpString As String) As LongPtr
Private Declare PtrSafe Function SendMessageTimeout Lib "user32" _
    Alias "SendMessageTimeoutA" (ByVal hWnd As LongPtr, ByVal msg As LongPtr, _
    ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal fuFlags As LongPtr, _
    ByVal uTimeout As LongPtr, ByRef lpdwResult As LongPtr) As LongPtr
Private Declare PtrSafe Function IIDFromString Lib "ole32" ( _
    ByRef lpsz As Any, ByRef lpiid As Any) As Long
Private Declare PtrSafe Function ObjectFromLresult Lib "oleacc" ( _
    ByVal lResult As LongPtr, ByRef riid As Any, ByVal wParam As LongPtr, _
    ByRef ppvObject As Object) As LongPtr

' 任意要素の配下から "Internet Explorer_Server" を探して IHTMLDocument を得る。
' elemPtr は生の IUIAutomationElement ポインタ。
Public Function GetIES_From_NomalElement(ByVal elemPtr As LongPtr) As Object
    If elemPtr = 0 Then
        Debug.Print "GetIES_From_NomalElement: 対象要素が Null です。"
        Exit Function
    End If

    Dim iesPtr As LongPtr
    iesPtr = FindIESElement(elemPtr)
    If iesPtr = 0 Then
        Debug.Print "GetIES_From_NomalElement: Internet Explorer_Server が見つかりません。"
        Exit Function
    End If

    Set GetIES_From_NomalElement = GetIES_From_ieElement(iesPtr)
    uia_core.ComRelease iesPtr
End Function

' "Internet Explorer_Server" 要素そのものから IHTMLDocument を得る。
Public Function GetIES_From_ieElement(ByVal elemPtr As LongPtr) As Object
    If elemPtr = 0 Then Exit Function

    Dim ClassName As String
    ClassName = uia_core.GetPropertyValue(elemPtr, UIA_ClassNamePropertyId)
    If ClassName <> "Internet Explorer_Server" Then
        Debug.Print "GetIES_From_ieElement: クラス名が " & ClassName & " です。中止します。"
        Exit Function
    End If

    Dim hWnd As LongPtr
    hWnd = uia_core.GetPropertyValue(elemPtr, UIA_NativeWindowHandlePropertyId)
    Set GetIES_From_ieElement = GetHTMLDocumentFromIES(hWnd)
End Function

' 任意要素の配下の IES から IHTMLDocument を得る (GetIES_From_NomalElement と同義)。
Public Function GetIES(ByVal elemPtr As LongPtr) As Object
    Set GetIES = GetIES_From_NomalElement(elemPtr)
End Function

' ウィンドウハンドルから IHTMLDocument2 (Object) を取得する。
' WM_HTML_GETOBJECT メッセージ + ObjectFromLresult による定番手法。API だけで動く。
Public Function GetHTMLDocumentFromIES(ByVal hWnd As LongPtr) As Object
    Const SMTO_ABORTIFHUNG As Long = &H2
    Const IID_IHTMLDocument2 As String = "{332c4425-26cb-11d0-b483-00c04fd90119}"

    Dim msg As LongPtr, res As LongPtr
    Dim iid(0 To 3) As LongPtr
    Dim obj As Object

    If hWnd = 0 Then Exit Function

    msg = RegisterWindowMessage("WM_HTML_GETOBJECT")
    SendMessageTimeout hWnd, msg, 0, 0, SMTO_ABORTIFHUNG, 1000, res
    If res <> 0 Then
        IIDFromString ByVal StrPtr(IID_IHTMLDocument2), iid(0)
        If ObjectFromLresult(res, iid(0), 0, obj) = 0 Then
            Set GetHTMLDocumentFromIES = obj
        End If
    End If
End Function

' 要素の配下 (Subtree) から "Internet Explorer_Server" 要素を探す。
' 戻り値は AddRef 済みの要素ポインタ (呼び出し側で ComRelease)。見つからなければ 0。
Private Function FindIESElement(ByVal elemPtr As LongPtr) As LongPtr
    Dim cond As LongPtr, found As LongPtr
    cond = uia_core.CreatePropertyCondition(UIA_ClassNamePropertyId, "Internet Explorer_Server")
    uia_core.UiaCheck uia_core.UiaInvoke(elemPtr, 5, "LPP", TreeScope_Subtree, cond, VarPtr(found)), _
                      "IUIAutomationElement.FindFirst (IES)"
    uia_core.ComRelease cond
    FindIESElement = found
End Function
