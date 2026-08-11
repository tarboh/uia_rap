VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "uia_c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

#Const DebugMode = 1

Private uia As New CUIAutomation

Public cnd As IUIAutomationCondition 'クラスの既定メンバー　"Attribute cnd.VB_VarUserMemId = 0"を仕込む
Attribute cnd.VB_VarUserMemId = 0

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
'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'コンディションの設定
'CreateNewInstanceがfalseだと、現在のインスタンスのcndを加工する形になる

'内部処理用の関数
Private Function SetCondition(Style As CndSetType, NewCnd As IUIAutomationCondition) As IUIAutomationCondition

    Select Case Style
        Case CndSetType.AsNew
            Set SetCondition = NewCnd
        Case CndSetType.AsAnd
            If cnd Is Nothing Then
                Set SetCondition = NewCnd
            Else
                Set SetCondition = uia.CreateAndCondition(cnd, NewCnd)
            End If
        Case CndSetType.AsOr
            If cnd Is Nothing Then
                Set SetCondition = NewCnd
            Else
                Set SetCondition = uia.CreateOrCondition(cnd, NewCnd)
            End If
    End Select
    
End Function



Public Function Name4_Full(Name_ As String, Optional MatchType As EnumMatchType = FullMatch, Optional Style As CndSetType = CndSetType.AsAnd, Optional CreateNewInstance As Boolean = True) As uia_c
    
    Dim cnd_ As IUIAutomationCondition
    Set cnd_ = uia.CreatePropertyConditionEx(UIA_NamePropertyId, Name_, MatchType)
    
    Dim Target As uia_c
    
    If CreateNewInstance Then
        Set Target = New uia_c
    Else
        Set Target = Me
    End If
    
    Set Target.cnd = SetCondition(Style, cnd_)
    
    Set Name4_Full = Target
    
End Function

Public Function Name1_Sub(Name_ As String, Optional MatchType As EnumMatchType = EnumMatchType.Substring_IgnorCase, Optional Style As CndSetType = CndSetType.AsAnd, Optional CreateNewInstance As Boolean = True) As uia_c
    
    Dim cnd_ As IUIAutomationCondition
    Set cnd_ = uia.CreatePropertyConditionEx(UIA_NamePropertyId, Name_, MatchType)
    
    Dim Target As uia_c
    
    If CreateNewInstance Then
        Set Target = New uia_c
    Else
        Set Target = Me
    End If
    
    Set Target.cnd = SetCondition(Style, cnd_)
    
    Set Name1_Sub = Target
    
End Function

Public Function Name2_Sub_Or(Name1 As String, Name2 As String, Optional MatchType As EnumMatchType = EnumMatchType.Substring_IgnorCase, Optional Style As CndSetType = CndSetType.AsAnd, Optional CreateNewInstance As Boolean = True) As uia_c
    
    Dim cnd_ As IUIAutomationCondition
    Set cnd_ = uia.CreateOrCondition( _
        uia.CreatePropertyConditionEx(UIA_NamePropertyId, Name1, MatchType), _
        uia.CreatePropertyConditionEx(UIA_NamePropertyId, Name2, MatchType))
    
    Dim Target As uia_c
    
    If CreateNewInstance Then
        Set Target = New uia_c
    Else
        Set Target = Me
    End If
    
    Set Target.cnd = SetCondition(Style, cnd_)
    
    Set Name2_Sub_Or = Target
    
End Function

Public Function Name3_Sub_And(Name1 As String, Name2 As String, Optional MatchType As EnumMatchType = EnumMatchType.Substring_IgnorCase, Optional Style As CndSetType = CndSetType.AsAnd, Optional CreateNewInstance As Boolean = True) As uia_c
    
    Dim cnd_ As IUIAutomationCondition
    Set cnd_ = uia.CreateAndCondition( _
        uia.CreatePropertyConditionEx(UIA_NamePropertyId, Name1, MatchType), _
        uia.CreatePropertyConditionEx(UIA_NamePropertyId, Name2, MatchType))
    
    Dim Target As uia_c
    
    If CreateNewInstance Then
        Set Target = New uia_c
    Else
        Set Target = Me
    End If
    
    Set Target.cnd = SetCondition(Style, cnd_)
    
    Set Name3_Sub_And = Target
    
End Function

Public Function Name5_Manual(Name_ As String, MatchType As EnumMatchType, Optional Style As CndSetType = CndSetType.AsAnd, Optional CreateNewInstance As Boolean = True) As uia_c
    
    Dim cnd_ As IUIAutomationCondition
    Set cnd_ = uia.CreatePropertyConditionEx(UIA_NamePropertyId, Name_, MatchType)
    
    Dim Target As uia_c
    
    If CreateNewInstance Then
        Set Target = New uia_c
    Else
        Set Target = Me
    End If
    
    Set Target.cnd = SetCondition(Style, cnd_)
    
    Set Name5_Manual = Target
    
End Function

Public Function Type_(ctrlType As EnumCtrlType, Optional Style As CndSetType = CndSetType.AsAnd, Optional CreateNewInstance As Boolean = True) As uia_c
    
    Dim cnd_ As IUIAutomationCondition
    Set cnd_ = uia.CreatePropertyCondition(UIA_ControlTypePropertyId, ctrlType)
        
    Dim Target As uia_c
    If CreateNewInstance Then
        Set Target = New uia_c
    Else
        Set Target = Me
    End If
    
    Set Target.cnd = SetCondition(Style, cnd_)
    
    Set Type_ = Target
    
End Function

Public Function LocalType(typeName As String, Optional Style As CndSetType = CndSetType.AsAnd, Optional CreateNewInstance As Boolean = True) As uia_c
    
    Dim cnd_ As IUIAutomationCondition
    Set cnd_ = uia.CreatePropertyCondition(UIA_LocalizedControlTypePropertyId, typeName)
        
    Dim Target As uia_c
    If CreateNewInstance Then
        Set Target = New uia_c
    Else
        Set Target = Me
    End If
    
    Set Target.cnd = SetCondition(Style, cnd_)
    
    Set LocalType = Target
    
End Function

Public Function ClsName(ClassName As String, Optional Style As CndSetType = CndSetType.AsAnd, Optional CreateNewInstance As Boolean = True) As uia_c
    
    Dim cnd_ As IUIAutomationCondition
    Set cnd_ = uia.CreatePropertyCondition(UIA_ClassNamePropertyId, ClassName)
        
    Dim Target As uia_c
    If CreateNewInstance Then
        Set Target = New uia_c
    Else
        Set Target = Me
    End If
    
    Set Target.cnd = SetCondition(Style, cnd_)
    
    Set ClsName = Target
    
End Function

Public Function AutomationId(AutomationId_ As String, Optional Style As CndSetType = CndSetType.AsAnd, Optional CreateNewInstance As Boolean = True) As uia_c
    
    Dim cnd_ As IUIAutomationCondition
    Set cnd_ = uia.CreatePropertyCondition(UIA_AutomationIdPropertyId, AutomationId_)
        
    Dim Target As uia_c
    If CreateNewInstance Then
        Set Target = New uia_c
    Else
        Set Target = Me
    End If
    
    Set Target.cnd = SetCondition(Style, cnd_)
    
    Set AutomationId = Target
    
End Function

Public Function True_(Optional Style As CndSetType = CndSetType.AsNew, Optional CreateNewInstance As Boolean = True) As uia_c
    
    Dim cnd_ As IUIAutomationCondition
    Set cnd_ = uia.CreateTrueCondition
        
    Dim Target As uia_c
    If CreateNewInstance Then
        Set Target = New uia_c
    Else
        Set Target = Me
    End If
    
    Set Target.cnd = SetCondition(Style, cnd_)
    
    Set True_ = Target
    
End Function

Public Function False_(Optional Style As CndSetType = CndSetType.AsNew, Optional CreateNewInstance As Boolean = True) As uia_c
    
    Dim cnd_ As IUIAutomationCondition
    Set cnd_ = uia.CreateFalseCondition
        
    Dim Target As uia_c
    If CreateNewInstance Then
        Set Target = New uia_c
    Else
        Set Target = Me
    End If
    
    Set Target.cnd = SetCondition(Style, cnd_)
    
    Set False_ = Target
    
End Function

Public Function Not_(cnd1 As uia_c, Optional CreateNewInstance As Boolean = True) As uia_c
    
    Dim cnd_ As IUIAutomationCondition
    If Alert_InstanceIsAvailable(cnd1, "AddDirect_Not") Then Exit Function
    If Alert_ConditionAvailable(cnd1.cnd, "AddDirect_Not") Then Exit Function
    
    Set cnd_ = uia.CreateNotCondition(cnd1.cnd)
    
    Dim Target As uia_c
    
    If CreateNewInstance Then
        Set Target = New uia_c
    Else
        Set Target = Me
    End If
    
    Set Target.cnd = uia.CreateNotCondition(cnd1.cnd)
    
    Set Not_ = Target
    
End Function


'Public Function Set_Add_And(cnd1 As uia_c, Optional CreateNewInstance As Boolean = False) As uia_c
'
'    Dim cnd_ As IUIAutomationCondition
'    If Alert_InstanceIsAvailable(cnd1, "AddDirect_And") Then Exit Function
'    If cnd Is Nothing Then
'        Set cnd_ = cnd1.cnd
'    Else
'        Set cnd_ = uia.CreateAndCondition(cnd, cnd1.cnd)
'    End If
'
'    If CreateNewInstance Then
'        Dim c As New uia_c
'        Set c.cnd = cnd_
'        Set Set_Add_And = c
'    Else
'        Set Me.cnd = cnd_
'        Set Set_Add_And = Me
'    End If
'
'End Function
'
'Public Function Set_Add_Or(cnd1 As uia_c, Optional CreateNewInstance As Boolean = False) As uia_c
'
'    Dim cnd_ As IUIAutomationCondition
'    If Alert_InstanceIsAvailable(cnd1, "AddDirect_Or") Then Exit Function
'    If cnd Is Nothing Then
'        Set cnd_ = cnd1.cnd
'    Else
'        Set cnd_ = uia.CreateOrCondition(cnd, cnd1.cnd)
'    End If
'
'    If CreateNewInstance Then
'        Dim c As New uia_c
'        Set c.cnd = cnd_
'        Set Set_Add_Or = c
'    Else
'        Set Me.cnd = cnd_
'        Set Set_Add_Or = Me
'    End If
'
'End Function




'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'デバッグ用

Function Alert_InstanceIsAvailable(c As uia_c, Optional Name As String) As Boolean
    If c Is Nothing Then
        #If DebugMode = 1 Then
            Debug.Print "uia_c Is Nothig! " & Name
        #End If
        Alert_InstanceIsAvailable = True
    End If
End Function

Function Alert_ConditionAvailable(cnd As IUIAutomationCondition, Name As String) As Boolean
    If cnd Is Nothing Then
        #If DebugMode = 1 Then
            Debug.Print "cnd Is Nothig! " & Name
        #End If
        Alert_ConditionAvailable = True
    End If
End Function

Function Alert_CtrlTypeName(ctrlType As Long)

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
    End Select

    Alert_CtrlTypeName = res

    #If DebugMode = 1 Then
        Debug.Print ctrlType & ":" & res
    #End If
    
End Function

'主要なクラス名をEnumで選択して返すサボり関数
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
