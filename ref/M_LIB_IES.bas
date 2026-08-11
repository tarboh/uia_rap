Attribute VB_Name = "M_LIB_IES"
Option Explicit

Private Declare PtrSafe Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As LongPtr
Private Declare PtrSafe Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As LongPtr, ByVal msg As LongPtr, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal fuFlags As LongPtr, ByVal uTimeout As LongPtr, lpdwResult As LongPtr) As LongPtr
Private Declare PtrSafe Function IIDFromString Lib "ole32" (lpsz As Any, lpiid As Any) As Long
Private Declare PtrSafe Function ObjectFromLresult Lib "oleacc" (ByVal lResult As LongPtr, riid As Any, ByVal wParam As LongPtr, ppvObject As Object) As LongPtr
Private hIES  As LongPtr

'IUIAutomationElementからInternet Explorer_Serverクラスのエレメントを検索し、IHTMLDocumentを取得する処理
Public Function GetIES_From_NomalElement(ByVal elem As IUIAutomationElement) As IHTMLDocument
    
    If elem Is Nothing Then
        
        Debug.Print "GetIES_From_NomalElement 実行エラー" & vbCrLf & _
                    "処理対象のエレメントが Nothing でした。処理を中止します。"
        Exit Function
        
    End If
    
    Dim uia As New CUIAutomation
    
    Dim elem_IES As IUIAutomationElement
    Set elem_IES = elem.FindFirst(TreeScope_Subtree, uia.CreatePropertyCondition(UIA_ClassNamePropertyId, "Internet Explorer_Server"))
    
    If Not elem_IES Is Nothing Then
        
        Set GetIES_From_NomalElement = GetIES_From_ieElement(elem_IES)
    
    Else
        
        Debug.Print "GetIES_From_NomalElement 実行エラー" & vbCrLf & _
                    "処理対象のエレメント内に Internet Explorer_Server クラスのエレメントが見つかりませんでした。処理を中止します。"
    
        Exit Function
        
    End If
    
End Function

'Internet Explorer_ServerクラスのIUIautomationElementからIHTMLDocumentを取得する処理
Public Function GetIES_From_ieElement(ByVal elem As IUIAutomationElement) As IHTMLDocument
        
    If Not elem Is Nothing Then
        
        Dim ClassName As String
        ClassName = elem.GetCurrentPropertyValue(UIA_ClassNamePropertyId)
        
        If ClassName <> "Internet Explorer_Server" Then
            Debug.Print "GetIES_From_ieElement 実行エラー" & vbCrLf & _
                        "処理対象エレメントのクラス名が " & ClassName & " でした。処理を中止します。"
            Exit Function
        End If
        
        Dim hWnd As LongPtr
        hWnd = elem.GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId)
        
        Set GetIES_From_ieElement = GetHTMLDocumentFromIES(hWnd)
        
    End If
    
End Function

'IUIautomationElementからIHTMLDocumentを取得する処理
Public Function GetIES(ByVal elem As IUIAutomationElement) As IHTMLDocument
    
    Dim uia As New CUIAutomation
    
    Dim elem_IES As IUIAutomationElement
    Set elem_IES = elem.FindFirst(TreeScope_Subtree, uia.CreatePropertyCondition(UIA_ClassNamePropertyId, "Internet Explorer_Server"))
    
    If Not elem_IES Is Nothing Then
        
        Dim hWnd As LongPtr
        hWnd = elem.GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId)
        
        Set GetIES = GetHTMLDocumentFromIES(hWnd)
        
    End If
    
End Function


'ウィンドウハンドルからIHTMLDocumentを取得する処理
Public Function GetHTMLDocumentFromIES(ByVal hWnd As LongPtr) As Object
    
    Dim msg As LongPtr, res As LongPtr
    Dim iid(0 To 3) As LongPtr
    Dim ret As Object, obj As Object
    Const SMTO_ABORTIFHUNG = &H2
  
    Const IID_IHTMLDocument1 = "{626FC520-A41E-11cf-A731-00A0C9082637}"
    Const IID_IHTMLDocument2 = "{332c4425-26cb-11d0-b483-00c04fd90119}"
    Const IID_IHTMLDocument3 = "{3050f485-98b5-11cf-bb82-00aa00bdce0b}"
    Const IID_IHTMLDocument4 = "{3050f69a-98b5-11cf-bb82-00aa00bdce0b}"
    Const IID_IHTMLDocument5 = "{3050f80c-98b5-11cf-bb82-00aa00bdce0b}"
    Const IID_IHTMLDocument6 = "{30510417-98b5-11cf-bb82-00aa00bdce0b}"
    Const IID_IHTMLDocument7 = "{305104b8-98b5-11cf-bb82-00aa00bdce0b}"
    Const IID_IHTMLDocument8 = "{305107d0-98b5-11cf-bb82-00aa00bdce0b}"
   
    Set ret = Nothing '初期化
    msg = RegisterWindowMessage("WM_HTML_GETOBJECT")
    SendMessageTimeout hWnd, msg, 0, 0, SMTO_ABORTIFHUNG, 1000, res
    If res Then
      IIDFromString StrPtr(IID_IHTMLDocument2), iid(0)
      If ObjectFromLresult(res, iid(0), 0, obj) = 0 Then Set ret = obj
    End If
    Set GetHTMLDocumentFromIES = ret
    
End Function
