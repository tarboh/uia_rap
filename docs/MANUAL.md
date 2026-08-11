# uia_rap メソッドリファレンス

[uia_rap](https://github.com/tarboh/uia_rap) の全公開メソッドの一覧です。`ref/`（参照設定あり）と `noref/`（参照設定なし）で **API は共通**なので、どちらでも同じように使えます。

- `uia_e` … 要素（Element）
- `uia_c` … 検索条件（Condition）
- `uia_t` … テキスト範囲（TextRange）
- `uia_Factory` … `e()` / `c()` / `t()` の生成入口
- `M_LIB_IES` … IE モードの HTML DOM 取得

各表の **「対応する UI Automation」列** は、そのラッパが内部で呼んでいる素の UI Automation メソッドです。UIA のドキュメントを引くときや、参照設定あり版で自前に書くときの手がかりにしてください（`IUIAutomation` = ルートオブジェクト、`IUIAutomationElement` = 要素、`IUIAutomationXxxPattern` = 各コントロールパターン）。

---

## 全体像

`uia_Factory` の関数がチェーンの起点です。呼ぶたびに新しいインスタンスを返します。

| 関数 | 戻り | 説明 |
|---|---|---|
| `e()` | `uia_e` | 空の要素インスタンス。`e.getRoot` などの起点 |
| `c()` | `uia_c` | 空の条件インスタンス。`c.Type_(Button)` などの起点 |
| `t()` | `uia_t` | 空のテキスト範囲インスタンス |

基本の流れは「**要素を取得 → 条件で検索 → プロパティ取得 or 操作**」です。

```vb
' 電卓ウィンドウを子から探し、その中の「7」ボタンを押す
e.getRoot.ffChildren(c.Name1_Sub("電卓")) _
         .ffDescendants(c.Type_(Button).Name4_Full("7")).ptInvoke
```

> **失敗時の挙動**：要素の検索系（`ff*`/`fa*`/`tw*` など）は、見つからなくても `Nothing` ではなく「中身が空の `uia_e`」を返します。そのままチェーンを続けられますが、確実に取れたか確かめたいときは `.prName` が空か、内部ポインタ `.elem = 0` かで判定できます。プロパティ（`pr*`）やパターン操作（`pt*`）は、要素が空なら何もしません。

---

## uia_c ― 検索条件

条件はチェーンで積み上げます。共通の省略可能引数：

- `Style`（`CndSetType`）… 既存条件との結合方法。既定 `AsAnd`。`AsOr` / `AsNew` も可
- `CreateNewInstance`（Boolean）… 既定 `True`（新しいインスタンスを返す）。`False` だと自身を書き換える

| メソッド | 説明 | 対応する UI Automation |
|---|---|---|
| `Name4_Full(name)` | 名前が**完全一致** | `IUIAutomation.CreatePropertyConditionEx(UIA_NamePropertyId, …)` |
| `Name1_Sub(name)` | 名前に**部分一致**（大小無視が既定） | 同上（flags に部分一致・大小無視） |
| `Name2_Sub_Or(a, b)` | 名前が a **または** b | `CreateOrCondition(CreatePropertyConditionEx ×2)` |
| `Name3_Sub_And(a, b)` | 名前が a **かつ** b | `CreateAndCondition(CreatePropertyConditionEx ×2)` |
| `Name5_Manual(name, matchType)` | 一致方法を明示指定 | `CreatePropertyConditionEx(UIA_NamePropertyId, …)` |
| `Type_(ctrlType)` | コントロール種別で絞る | `CreatePropertyCondition(UIA_ControlTypePropertyId, …)` |
| `LocalType(name)` | ローカライズ種別名で絞る | `CreatePropertyCondition(UIA_LocalizedControlTypePropertyId, …)` |
| `ClsName(className)` | クラス名で絞る | `CreatePropertyCondition(UIA_ClassNamePropertyId, …)` |
| `AutomationId(id)` | AutomationId で絞る | `CreatePropertyCondition(UIA_AutomationIdPropertyId, …)` |
| `True_()` | 常に真（全要素にマッチ） | `IUIAutomation.CreateTrueCondition` |
| `False_()` | 常に偽 | `IUIAutomation.CreateFalseCondition` |
| `Not_(cnd1)` | 条件の否定 | `IUIAutomation.CreateNotCondition` |
| `CNMS(id)` | よく使うクラス名を Enum で返す補助（下表） | ―（文字列を返すだけ） |

チェーンで積むときの AND / OR 結合は、内部で `IUIAutomation.CreateAndCondition` / `CreateOrCondition` を呼んでいます。`cnd`（既定メンバー）が組み立てた条件の生ポインタです。

```vb
Set cond = c.Type_(Button).Name4_Full("保存")     ' Button AND 名前="保存"
Set cond = c.Name2_Sub_Or("はい", "OK")           ' 名前が "はい" または "OK"
Set cond = c.ClsName("Edit").AutomationId("15")    ' クラス名 AND AutomationId
```

### CNMS が返すクラス名（`eCNMS`）

| メンバー | 返す文字列 |
|---|---|
| `Edge_Top` | `Chrome_WidgetWin_1` |
| `Edge_AddresBar` | `OmniboxViewViews` |
| `Edge_TabBase` | `EdgeTabContainerImpl` |
| `Excel_Top` | `XLMAIN` |
| `Excel_SheetBase` | `ExcelGrid` |
| `Excel_SheetGrid` | `XLSpreadsheetGrid` |
| `IE_IEServer` | `Internet Explorer_Server` |

---

## uia_e ― 要素

### 取得

| メソッド | 説明 | 対応する UI Automation |
|---|---|---|
| `getRoot()` | デスクトップ（ルート要素） | `IUIAutomation.GetRootElement` |
| `getFocus()` | フォーカス中の要素 | `IUIAutomation.GetFocusedElement` |
| `getHandle(hwnd)` | HWND から要素 | `IUIAutomation.ElementFromHandle` |
| `getPoint(pt)` / `GetFromPoint(x, y)` | 座標から要素 | `IUIAutomation.ElementFromPoint` |
| `GetFromCursor()` | カーソル直下の要素 | `GetCursorPos` + `ElementFromPoint` |

### 検索

| メソッド | 説明 | 対応する UI Automation |
|---|---|---|
| `ffChildren(c, …)` | **子**から最初の1件 | `IUIAutomationElement.FindFirst(TreeScope_Children, …)` |
| `ffDescendants(c, …, [direction])` | **子孫**から最初の1件 | `FindFirst(TreeScope_Descendants, …)` / direction 指定時は `IUIAutomationElement9.FindFirstWithOptions` |
| `faChildren([c], …)` | **子**を全件 | `IUIAutomationElement.FindAll(TreeScope_Children, …)` |
| `faDescendants(c, …, [direction])` | **子孫**を全件 | `FindAll(TreeScope_Descendants, …)` / direction 指定時は `IUIAutomationElement9.FindAllWithOptions` |

- `RetryCount` / `RetrySleepTime`：見つからないとき指定回数リトライ（描画待ちに有効）
- `direction`（`EnumFindDirection`）：`Default` / `PostOrder` / `LastToFirst`

```vb
Set el = parent.ffDescendants(c.Type_(Button).Name4_Full("OK"), RetryCount:=5)

Dim arr As uia_e, i As Long
Set arr = parent.faChildren(c.Type_(ListItem))
For i = 0 To arr.Array_Length - 1
    Debug.Print arr.Array_GetItemByIndex(i).prName
Next
```

### TreeWalker（木をたどる）

| メソッド | 説明 | 対応する UI Automation |
|---|---|---|
| `tw_Set(twType, [cnd])` | 使う walker を設定 | `IUIAutomation.get_ControlViewWalker` / `get_ContentViewWalker` / `get_RawViewWalker` / `CreateTreeWalker` |
| `twParent()` | 親へ | `IUIAutomationTreeWalker.GetParentElement` |
| `twFirstChild()` | 最初の子へ | `GetFirstChildElement` |
| `twLastChild()` | 最後の子へ | `GetLastChildElement` |
| `twNext()` | 次の兄弟へ | `GetNextSiblingElement` |
| `twPrev()` | 前の兄弟へ | `GetPreviousSiblingElement` |

`tw_Set` を呼ばずに `tw*` を使うと `ControlView` が既定で使われます。

### プロパティ（`pr*`）

| プロパティ | 型 | 対応する UI Automation |
|---|---|---|
| `prName` | String | `IUIAutomationElement.get_CurrentName` |
| `prClsName` | String | `get_CurrentClassName` |
| `prCtrlType` | Long | `get_CurrentControlType` |
| `prShortCtrlTypeName` | String | `get_CurrentControlType`（値を短い名前に変換） |
| `prLocalCtrlType` | String | `get_CurrentLocalizedControlType` |
| `prValue` | String | `GetCurrentPropertyValue(UIA_ValueValuePropertyId)` |
| `prHwnd` | LongPtr | `GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId)` |
| `prToggleState` | Long | `GetCurrentPropertyValue(UIA_ToggleToggleStatePropertyId)` |
| `prExpandCollapseState` | Long | `GetCurrentPropertyValue(UIA_ExpandCollapseExpandCollapseStatePropertyId)` |
| `prScrollState_X` / `_Y` | Double | `GetCurrentPropertyValue(UIA_Scroll…ScrollPercentPropertyId)` |
| `prRect` | `tagRECT` | `get_CurrentBoundingRectangle`（**{Left, Top, Right, Bottom}**） |
| `prRectLeft` / `Top` / `Right` / `Bottom` / `prRectCenter` | Long / `tagPOINT` | 同上（`get_CurrentBoundingRectangle`） |
| `prPtnAv_Invoke` | Boolean | `GetCurrentPropertyValue(UIA_IsInvokePatternAvailablePropertyId)` |

```vb
Debug.Print el.prName & " [" & el.prLocalCtrlType & "] {" & el.prClsName & "}"
Debug.Print "中心: " & el.prRectCenter.X & "," & el.prRectCenter.Y
```

### 操作（`pt*` ほか）

パターン系は内部で `IUIAutomationElement.GetCurrentPattern(パターンID)` を呼んでから、そのパターンのメソッドを実行します。

| メソッド | 対応する UI Automation |
|---|---|
| `ptInvoke()` | `IUIAutomationInvokePattern.Invoke` |
| `ptSetValue(value)` | `IUIAutomationValuePattern.SetValue` |
| `ptToggle()` | `IUIAutomationTogglePattern.Toggle` |
| `ptExpand()` / `ptCollapse()` | `IUIAutomationExpandCollapsePattern.Expand` / `Collapse` |
| `ptScroll([xPct], [yPct])` | `IUIAutomationScrollPattern.SetScrollPercent`（省略時は `get_CurrentHorizontal/VerticalScrollPercent`） |
| `ptScrollIntoView()` | `IUIAutomationScrollItemPattern.ScrollIntoView` |
| `ptSelectionItem_Select()` | `IUIAutomationSelectionItemPattern.Select` |
| `ptSelection_GetSelection()` | `IUIAutomationSelectionPattern.GetCurrentSelection` |
| `ptGridGetItem(row, col)` | `IUIAutomationGridPattern.GetItem` |
| `ptWindowClose()` | `IUIAutomationWindowPattern.Close` |
| `ptGetTextRange()` | `IUIAutomationTextPattern(2)/TextChild/TextEdit` の `get_DocumentRange` / `get_TextRange` |
| `SetFocus()` | `IUIAutomationElement.SetFocus` |
| `ShowContextmenu()` | `IUIAutomationElement3.ShowContextMenu` |
| `CompareTo(target)` | `IUIAutomation.CompareElements` |

### 配列（`fa*` の結果を扱う）

| メソッド | 説明 | 対応する UI Automation |
|---|---|---|
| `Array_Length` | 要素配列の件数 | `IUIAutomationElementArray.get_Length` |
| `Array_GetItemByIndex(i)` | i 番目の要素 | `IUIAutomationElementArray.GetElement` |
| `ElemArray_Col` | 全要素を `Collection` で取得 | 同上（GetElement を回す） |
| `Child_SetElemArray([c])` | 子要素を条件で取り込む | `FindAll(TreeScope_Children, …)`（条件省略時 `CreateTrueCondition`） |
| `Child_GetItem(c)` | 各子の配下を条件で探し最初の子を返す | `FindAll` + `FindFirst` |

### Edge 操作

現在の要素が Edge のトップウィンドウ（`Chrome_WidgetWin_1` / 名前が `*Edge`）である前提のものが多いです。`AdditionalCondition` は文字列（名前の部分一致）か `uia_c` を渡せます。内部的には `ffChildren` / `faChildren` / `ffDescendants` と条件（`ClsName` 等）の組み合わせで、UI Automation の `FindFirst` / `FindAll` を呼んでいます。

| メソッド | 説明 |
|---|---|
| `EdgeGetTopWindow([cond])` | Edge のトップウィンドウを1件取得 |
| `EdgeGetTopWindows([cond])` | 同・全件 |
| `EdgeGetTabItems([cond])` | タブアイテム群を取得 |
| `EdgeChangeTab(cond)` | 条件に合うタブへ切り替え |
| `EdgeCreateNewTab()` | 新しいタブを開く |
| `EdgeGetAddresBar()` | アドレスバー要素を取得 |
| `EdgeSetAddress(url, [enter])` | アドレスバーに入力（`enter:=True` で Enter 送信） |

### IE モードの HTML DOM

| メソッド | 戻り | 説明 |
|---|---|---|
| `IES_GetIHTMLDocument_FromTopWindow()` | Object | 配下の `Internet Explorer_Server` から `IHTMLDocument2` |
| `IES_GetIHTMLDocument_FromIESElement()` | Object | 現在の要素が IES 要素そのものの場合 |

（UI Automation としては、配下から `Internet Explorer_Server` を `FindFirst` し、その `NativeWindowHandle` を使う。詳細は下の `M_LIB_IES` 参照）

```vb
Dim doc As Object
Set doc = edgeTab.IES_GetIHTMLDocument_FromTopWindow()
Debug.Print doc.title
doc.getElementById("foo").Click
```

### その他 / 応用

| メソッド | 説明 |
|---|---|
| `GetTopWindow()` | 祖先をたどってトップウィンドウ（デスクトップの1つ下）を取得（`twParent` を反復） |
| `ptValue_DLFile([folder], [savedPath], [overwrite])` | ValuePattern の URL を `URLDownloadToFile` で保存 |
| `GetCtrlTypeName(id)` | 種別 ID → 定数名の文字列 |
| `GetInfo([verbose])` | プロパティ/パターン一覧を表示。内部で `IUIAutomation.PollForPotentialSupportedProperties` / `…Patterns` |
| `test()` | カーソル直下の要素を CTRL で拾って情報表示し続ける（ALT で終了） |
| `Sleep_(ms)` | 指定ミリ秒待つ（`kernel32.Sleep`） |

---

## uia_t ― テキスト範囲

`uia_e.ptGetTextRange()` で取得します。`Returntype`（`EnumtRngSetType`）は `AsSelfInstance`（自身を書き換え）/ `AsnewInstance`（新規）。

| メソッド | 説明 | 対応する UI Automation |
|---|---|---|
| `GetText([maxLength])` | 範囲のテキスト（-1 で全文） | `IUIAutomationTextRange.GetText` |
| `Find(returnType, text, …)` | テキストを検索して範囲を移動 | `IUIAutomationTextRange.FindText` |
| `Expand(returnType, unit)` | 単位境界まで範囲を広げる | `IUIAutomationTextRange.ExpandToEnclosingUnit` |
| `Move(returnType, unit, count)` | 範囲を移動 | `IUIAutomationTextRange.MoveEndpointByUnit`（始点・終点） |
| `ScrollIntoView(alignToTop)` | 範囲を表示位置へ | `IUIAutomationTextRange.ScrollIntoView` |
| `Select_()` | 範囲を選択 | `IUIAutomationTextRange.Select` |

`EnumTextUnitSize`：`a_Character` / `b_Format` / `c_Word` / `d_Line` / `e_Paragraph` / `f_Page` / `g_Document`

---

## M_LIB_IES ― IE モードの HTML DOM（下回り）

通常は `uia_e.IES_*` 経由で使いますが、直接も呼べます。要素は生ポインタ、戻りは `Object`（`IHTMLDocument2`）。

| 関数 | 説明 | 使う API |
|---|---|---|
| `GetIES_From_NomalElement(elemPtr)` | 配下の IES から取得 | `IUIAutomationElement.FindFirst`（クラス名条件）→ 下の2つ |
| `GetIES_From_ieElement(elemPtr)` | IES 要素そのものから取得 | `GetCurrentPropertyValue(ClassName / NativeWindowHandle)` |
| `GetIES(elemPtr)` | 上の別名 | 同上 |
| `GetHTMLDocumentFromIES(hwnd)` | HWND から直接取得 | `WM_HTML_GETOBJECT` + `ObjectFromLresult`（UIA ではなく Win32/oleacc） |

---

## Enum 早見表

**EnumCtrlType**（`Type_` に渡す種別。VBA 予約語と衝突するものは末尾に `_`）
`AppBar` `Button` `Calendar` `CheckBox` `ComboBox` `Custom` `DataGrid` `DataItem` `Document` `Edit` `Group` `Header` `HeaderItem` `Hyperlink` `Image` `List` `ListItem` `MenuBar` `Menu` `MenuItem` `Pane` `ProgressBar` `RadioButton` `ScrollBar` `SemanticZoom` `Separator` `Slider` `Spinner` `SplitButton` `StatusBar` `Tab_` `TabItem` `Table` `Text_` `Thumb` `TitleBar` `Toolbar` `ToolTip` `Tree` `TreeItem` `Window_`

**EnumMatchType**（名前の一致方法）
`FullMatch`（完全一致）/ `Substring`（部分一致）/ `IgnorCase`（大小無視）/ `Substring_IgnorCase`（部分一致・大小無視）

**CndSetType**（条件の結合）
`AsAnd`（かつ）/ `AsOr`（または）/ `AsNew`（置換）

**EnumFindDirection**（子孫検索の走査順）
`Default` / `PostOrder` / `LastToFirst`

**EnumTreeWakerType**（TreeWalker の種類）
`ControlView` / `ContentView` / `RawView` / `ByCondition`

---

## `UIA_ElementFromPoint.bas` について（`ref/` のみ）

`ref/` 版には `UIA_ElementFromPoint.bas` という補助モジュールが含まれます。**「座標から要素を取得する」処理だけ `DispCallFunc` で書かれている**のには理由があります。

UI Automation の `ElementFromPoint` は、引数の `POINT` 構造体を**値渡し**で受け取ります。ところが **VBA（VBE）は構造体（ユーザー定義型）を `ByVal` で渡すことを許していません**。参照設定を入れて `uia.ElementFromPoint(pt)` と普通に書いても、次のコンパイルエラーで弾かれます。

```
コンパイル エラー:
ユーザー定義型を ByVal で渡すことはできません。
```

つまりこのメソッドだけは通常の呼び出しが**そもそもできません**。そこで `DispCallFunc` で VBA のコンパイラを迂回し、構造体をバイナリレベルで渡しています。（さらに 64bit では 8 バイトの `POINT` を 1 つの 64bit 値に詰めて渡す必要があり、`noref/` 版ではこれを算術で組み立てて `CopyMemory` なしで実装しています。）

| 関数 | 戻り | 説明 |
|---|---|---|
| `ElementFromPoint(pt)` | `IUIAutomationElement` | 座標（`PointAPI` 構造体）から要素 |
| `ElementFromCursor()` | `IUIAutomationElement` | マウスカーソル直下の要素 |
| `ElementFromPoint動作サンプル()` | ― | カーソルを追いながら要素名を出す動作サンプル |

**`noref/` 版では、この処理は `uia_core` に統合済み**です（`uia_core.ElementFromPoint` として実装）。利用者は `uia_e.GetFromPoint(x, y)` / `uia_e.GetFromCursor()` から使えばよく、`UIA_ElementFromPoint.bas` を別途インポートする必要はありません。まとめると：

- **`ref/`** … 座標取得は `UIA_ElementFromPoint.bas`（別モジュール）が担当
- **`noref/`** … `uia_core` に統合。`GetFromPoint` / `GetFromCursor` から利用

---

## 補足

- `ref/`（参照あり）と `noref/`（参照なし）で API は同じですが、`noref/` は内部で `elem` / `cnd` / `tRng` を生ポインタ（`LongPtr`）として持ち、参照カウントを各クラスが管理しています。通常は意識せずに使えます。
- 「対応する UI Automation」に出てくる `IUIAutomation` / `IUIAutomationElement` などのインターフェースやメソッドは、Microsoft の [UI Automation クライアント API](https://learn.microsoft.com/windows/win32/api/uiautomationclient/) のドキュメントで引けます。
- 詳しい導入手順とサンプルは [README](https://github.com/tarboh/uia_rap) を参照してください。
