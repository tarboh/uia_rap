VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "uia_c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==============================================================================
' uia_c (参照設定不要版)  ―  検索条件のチェーンビルダ
'
' 参照設定アリ版との違い:
'   * New CUIAutomation を持たず、uia_core の共有 IUIAutomation を使う
'   * 既定メンバー cnd は IUIAutomationCondition オブジェクトではなく、
'     生の条件ポインタ (LongPtr)。参照カウントは本クラスが管理する。
'
' 使い方は参照アリ版と同じ:
'   Set cond = c.Type_(Button).Name4_Full("保存")
'==============================================================================
Option Explicit

#Const DebugMode = 1

' 条件ポインタ。クラスの既定メンバー (UserMemId = 0)。参照カウントは本クラスが持つ。
Private m_cnd As LongPtr

Public Dic_CtrlTypeName As Object

Public Enum CndSetType
    AsAnd
    AsOr
    AsNew
End Enum

Public Enum EnumMatchType
    FullMatch = PropertyConditionFlags_None
    Substring = PropertyConditionFlags_MatchSubstring
    IgnorCase = PropertyConditionFlags_IgnoreCase
    Substring_IgnorCase = PropertyConditionFlags_MatchSubstring + PropertyConditionFlags_IgnoreCase
End Enum

Public Enum EnumCtrlType
    AppBar = UIA_AppBarControlTypeId
    Button = UIA_ButtonControlTypeId
    Calendar = UIA_CalendarControlTypeId
    CheckBox = UIA_CheckBoxControlTypeId
    ComboBox = UIA_ComboBoxControlTypeId
    Custom = UIA_CustomControlTypeId
    DataGrid = UIA_DataGridControlTypeId
    DataItem = UIA_DataItemControlTypeId
    Document = UIA_DocumentControlTypeId
    Edit = UIA_EditControlTypeId
    Group = UIA_GroupControlTypeId
    Header = UIA_HeaderControlTypeId
    HeaderItem = UIA_HeaderItemControlTypeId
    Hyperlink = UIA_HyperlinkControlTypeId
    Image = UIA_ImageControlTypeId
    List = UIA_ListControlTypeId
    ListItem = UIA_ListItemControlTypeId
    MenuBar = UIA_MenuBarControlTypeId
    Menu = UIA_MenuControlTypeId
    MenuItem = UIA_MenuItemControlTypeId
    Pane = UIA_PaneControlTypeId
    ProgressBar = UIA_ProgressBarControlTypeId
    RadioButton = UIA_RadioButtonControlTypeId
    ScrollBar = UIA_ScrollBarControlTypeId
    SemanticZoom = UIA_SemanticZoomControlTypeId
    Separator = UIA_SeparatorControlTypeId
    Slider = UIA_SliderControlTypeId
    Spinner = UIA_SpinnerControlTypeId
    SplitButton = UIA_SplitButtonControlTypeId
    StatusBar = UIA_StatusBarControlTypeId
    Tab_ = UIA_TabControlTypeId
    TabItem = UIA_TabItemControlTypeId
    Table = UIA_TableControlTypeId
    Text_ = UIA_TextControlTypeId
    Thumb = UIA_ThumbControlTypeId
    TitleBar = UIA_TitleBarControlTypeId
    Toolbar = UIA_ToolBarControlTypeId
    ToolTip = UIA_ToolTipControlTypeId
    Tree = UIA_TreeControlTypeId
    TreeItem = UIA_TreeItemControlTypeId
    Window_ = UIA_WindowControlTypeId
End Enum

Public Enum eCNMS
    Edge_Top
    Edge_AddresBar
    Edge_TabBase
    Excel_Top
    Excel_SheetBase
    Excel_SheetGrid
    IE_IEServer
End Enum

'==============================================================================
' 既定メンバー cnd (生の条件ポインタ)
'==============================================================================

' 条件ポインタを取得する (所有権は移動しない)。クラスの既定メンバー。
Public Property Get cnd() As LongPtr
Attribute cnd.VB_UserMemId = 0
    cnd = m_cnd
End Property

' 条件ポインタを設定する。所有権を引き受け、古い条件は解放する。
Public Property Let cnd(ByVal newPtr As LongPtr)
    If m_cnd = newPtr Then Exit Property
    If m_cnd <> 0 Then uia_core.ComRelease m_cnd
    m_cnd = newPtr
End Property

Private Sub Class_Terminate()
    If m_cnd <> 0 Then uia_core.ComRelease m_cnd
End Sub

'==============================================================================
' 内部: Style に従って条件を合成する
'   newCnd の所有権は本関数が引き受け、戻り値として所有権を返す。
'==============================================================================
Private Function SetCondition(ByVal Style As CndSetType, ByVal newCnd As LongPtr) As LongPtr
    Select Case Style
        Case CndSetType.AsNew
            SetCondition = newCnd
        Case CndSetType.AsAnd
            If m_cnd = 0 Then
                SetCondition = newCnd
            Else
                SetCondition = uia_core.CreateAndCondition(m_cnd, newCnd)
                uia_core.ComRelease newCnd      ' 合成に取り込まれたので手放す
            End If
        Case CndSetType.AsOr
            If m_cnd = 0 Then
                SetCondition = newCnd
            Else
                SetCondition = uia_core.CreateOrCondition(m_cnd, newCnd)
                uia_core.ComRelease newCnd
            End If
    End Select
End Function

'==============================================================================
' 条件ビルダ
'==============================================================================

Public Function Name4_Full(Name_ As String, Optional MatchType As EnumMatchType = FullMatch, Optional Style As CndSetType = CndSetType.AsAnd, Optional CreateNewInstance As Boolean = True) As uia_c
    Dim cnd_ As LongPtr
    cnd_ = uia_core.CreatePropertyConditionEx(UIA_NamePropertyId, Name_, MatchType)
    Dim Target As uia_c
    If CreateNewInstance Then Set Target = New uia_c Else Set Target = Me
    Target.cnd = SetCondition(Style, cnd_)
    Set Name4_Full = Target
End Function

Public Function Name1_Sub(Name_ As String, Optional MatchType As EnumMatchType = EnumMatchType.Substring_IgnorCase, Optional Style As CndSetType = CndSetType.AsAnd, Optional CreateNewInstance As Boolean = True) As uia_c
    Dim cnd_ As LongPtr
    cnd_ = uia_core.CreatePropertyConditionEx(UIA_NamePropertyId, Name_, MatchType)
    Dim Target As uia_c
    If CreateNewInstance Then Set Target = New uia_c Else Set Target = Me
    Target.cnd = SetCondition(Style, cnd_)
    Set Name1_Sub = Target
End Function

Public Function Name2_Sub_Or(Name1 As String, Name2 As String, Optional MatchType As EnumMatchType = EnumMatchType.Substring_IgnorCase, Optional Style As CndSetType = CndSetType.AsAnd, Optional CreateNewInstance As Boolean = True) As uia_c
    Dim a As LongPtr, b As LongPtr, cnd_ As LongPtr
    a = uia_core.CreatePropertyConditionEx(UIA_NamePropertyId, Name1, MatchType)
    b = uia_core.CreatePropertyConditionEx(UIA_NamePropertyId, Name2, MatchType)
    cnd_ = uia_core.CreateOrCondition(a, b)
    uia_core.ComRelease a
    uia_core.ComRelease b
    Dim Target As uia_c
    If CreateNewInstance Then Set Target = New uia_c Else Set Target = Me
    Target.cnd = SetCondition(Style, cnd_)
    Set Name2_Sub_Or = Target
End Function

Public Function Name3_Sub_And(Name1 As String, Name2 As String, Optional MatchType As EnumMatchType = EnumMatchType.Substring_IgnorCase, Optional Style As CndSetType = CndSetType.AsAnd, Optional CreateNewInstance As Boolean = True) As uia_c
    Dim a As LongPtr, b As LongPtr, cnd_ As LongPtr
    a = uia_core.CreatePropertyConditionEx(UIA_NamePropertyId, Name1, MatchType)
    b = uia_core.CreatePropertyConditionEx(UIA_NamePropertyId, Name2, MatchType)
    cnd_ = uia_core.CreateAndCondition(a, b)
    uia_core.ComRelease a
    uia_core.ComRelease b
    Dim Target As uia_c
    If CreateNewInstance Then Set Target = New uia_c Else Set Target = Me
    Target.cnd = SetCondition(Style, cnd_)
    Set Name3_Sub_And = Target
End Function

Public Function Name5_Manual(Name_ As String, MatchType As EnumMatchType, Optional Style As CndSetType = CndSetType.AsAnd, Optional CreateNewInstance As Boolean = True) As uia_c
    Dim cnd_ As LongPtr
    cnd_ = uia_core.CreatePropertyConditionEx(UIA_NamePropertyId, Name_, MatchType)
    Dim Target As uia_c
    If CreateNewInstance Then Set Target = New uia_c Else Set Target = Me
    Target.cnd = SetCondition(Style, cnd_)
    Set Name5_Manual = Target
End Function

Public Function Type_(ctrlType As EnumCtrlType, Optional Style As CndSetType = CndSetType.AsAnd, Optional CreateNewInstance As Boolean = True) As uia_c
    Dim cnd_ As LongPtr
    cnd_ = uia_core.CreatePropertyCondition(UIA_ControlTypePropertyId, CLng(ctrlType))
    Dim Target As uia_c
    If CreateNewInstance Then Set Target = New uia_c Else Set Target = Me
    Target.cnd = SetCondition(Style, cnd_)
    Set Type_ = Target
End Function

Public Function LocalType(typeName As String, Optional Style As CndSetType = CndSetType.AsAnd, Optional CreateNewInstance As Boolean = True) As uia_c
    Dim cnd_ As LongPtr
    cnd_ = uia_core.CreatePropertyCondition(UIA_LocalizedControlTypePropertyId, typeName)
    Dim Target As uia_c
    If CreateNewInstance Then Set Target = New uia_c Else Set Target = Me
    Target.cnd = SetCondition(Style, cnd_)
    Set LocalType = Target
End Function

Public Function ClsName(ClassName As String, Optional Style As CndSetType = CndSetType.AsAnd, Optional CreateNewInstance As Boolean = True) As uia_c
    Dim cnd_ As LongPtr
    cnd_ = uia_core.CreatePropertyCondition(UIA_ClassNamePropertyId, ClassName)
    Dim Target As uia_c
    If CreateNewInstance Then Set Target = New uia_c Else Set Target = Me
    Target.cnd = SetCondition(Style, cnd_)
    Set ClsName = Target
End Function

Public Function AutomationId(AutomationId_ As String, Optional Style As CndSetType = CndSetType.AsAnd, Optional CreateNewInstance As Boolean = True) As uia_c
    Dim cnd_ As LongPtr
    cnd_ = uia_core.CreatePropertyCondition(UIA_AutomationIdPropertyId, AutomationId_)
    Dim Target As uia_c
    If CreateNewInstance Then Set Target = New uia_c Else Set Target = Me
    Target.cnd = SetCondition(Style, cnd_)
    Set AutomationId = Target
End Function

Public Function True_(Optional Style As CndSetType = CndSetType.AsNew, Optional CreateNewInstance As Boolean = True) As uia_c
    Dim cnd_ As LongPtr
    cnd_ = uia_core.CreateTrueCondition()
    Dim Target As uia_c
    If CreateNewInstance Then Set Target = New uia_c Else Set Target = Me
    Target.cnd = SetCondition(Style, cnd_)
    Set True_ = Target
End Function

Public Function False_(Optional Style As CndSetType = CndSetType.AsNew, Optional CreateNewInstance As Boolean = True) As uia_c
    Dim cnd_ As LongPtr
    cnd_ = uia_core.CreateFalseCondition()
    Dim Target As uia_c
    If CreateNewInstance Then Set Target = New uia_c Else Set Target = Me
    Target.cnd = SetCondition(Style, cnd_)
    Set False_ = Target
End Function

Public Function Not_(cnd1 As uia_c, Optional CreateNewInstance As Boolean = True) As uia_c
    If Alert_InstanceIsAvailable(cnd1, "Not_") Then Exit Function
    If Alert_ConditionAvailable(cnd1.cnd, "Not_") Then Exit Function

    Dim cnd_ As LongPtr
    cnd_ = uia_core.CreateNotCondition(cnd1.cnd)
    Dim Target As uia_c
    If CreateNewInstance Then Set Target = New uia_c Else Set Target = Me
    Target.cnd = cnd_
    Set Not_ = Target
End Function

'==============================================================================
' デバッグ / 補助
'==============================================================================

Function Alert_InstanceIsAvailable(c As uia_c, Optional Name As String) As Boolean
    If c Is Nothing Then
        #If DebugMode = 1 Then
            Debug.Print "uia_c Is Nothing! " & Name
        #End If
        Alert_InstanceIsAvailable = True
    End If
End Function

Function Alert_ConditionAvailable(ByVal cnd As LongPtr, Name As String) As Boolean
    If cnd = 0 Then
        #If DebugMode = 1 Then
            Debug.Print "cnd Is Null! " & Name
        #End If
        Alert_ConditionAvailable = True
    End If
End Function

' 主要なクラス名を Enum で選択して返すサボり関数
Public Function CNMS(clsNameId As eCNMS) As String
    Select Case clsNameId
        Case Edge_Top: CNMS = "Chrome_WidgetWin_1"
        Case Edge_AddresBar: CNMS = "OmniboxViewViews"
        Case Edge_TabBase: CNMS = "EdgeTabContainerImpl"
        Case Excel_Top: CNMS = "XLMAIN"
        Case Excel_SheetBase: CNMS = "ExcelGrid"
        Case Excel_SheetGrid: CNMS = "XLSpreadsheetGrid"
        Case IE_IEServer: CNMS = "Internet Explorer_Server"
    End Select
End Function
