# uia_rap メソッドリファレンス

[uia_rap](https://github.com/tarboh/uia_rap) の全公開メソッドの一覧です。`ref/`（参照設定あり）と `noref/`（参照設定なし）で **API は共通**なので、どちらでも同じように使えます。

- `uia_e` … 要素（Element）
- `uia_c` … 検索条件（Condition）
- `uia_t` … テキスト範囲（TextRange）
- `uia_Factory` … `e()` / `c()` / `t()` の生成入口
- `M_LIB_IES` … IE モードの HTML DOM 取得

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

| メソッド | 主な引数 | 説明 |
|---|---|---|
| `Name4_Full(name)` | name | 名前が**完全一致** |
| `Name1_Sub(name)` | name | 名前に**部分一致**（大小無視が既定） |
| `Name2_Sub_Or(a, b)` | a, b | 名前が a **または** b（部分一致） |
| `Name3_Sub_And(a, b)` | a, b | 名前が a **かつ** b（部分一致） |
| `Name5_Manual(name, matchType)` | name, matchType | 一致方法を明示指定 |
| `Type_(ctrlType)` | ctrlType（`EnumCtrlType`） | コントロール種別で絞る |
| `LocalType(name)` | name | ローカライズされた種別名で絞る |
| `ClsName(className)` | className | クラス名で絞る |
| `AutomationId(id)` | id | AutomationId で絞る |
| `True_()` | ― | 常に真（全要素にマッチ） |
| `False_()` | ― | 常に偽 |
| `Not_(cnd1)` | cnd1（`uia_c`） | 条件の否定 |
| `CNMS(id)` | id（`eCNMS`） | よく使うクラス名を Enum で返す補助（下表） |

`cnd`（既定メンバー）は組み立てた条件の生ポインタです。

```vb
' 種類=Button かつ 名前="保存"
Set cond = c.Type_(Button).Name4_Full("保存")

' 名前が "はい" または "OK"（部分一致）
Set cond = c.Name2_Sub_Or("はい", "OK")

' クラス名 AND AutomationId
Set cond = c.ClsName("Edit").AutomationId("15")
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

| メソッド | 引数 | 説明 |
|---|---|---|
| `getRoot()` | ― | デスクトップ（ルート要素） |
| `getFocus()` | ― | 現在フォーカスされている要素 |
| `getHandle(hwnd)` | ウィンドウハンドル | HWND から要素 |
| `getPoint(pt)` | `tagPOINT` | 座標（構造体）から要素 |
| `GetFromPoint(x, y)` | x, y | 座標（数値）から要素 |
| `GetFromCursor()` | ― | マウスカーソル直下の要素 |

### 検索

| メソッド | 引数 | 説明 |
|---|---|---|
| `ffChildren(c, [Retry], [Sleep])` | 条件, リトライ回数, 間隔ms | **子**から最初の1件 |
| `ffDescendants(c, [Retry], [Sleep], [direction])` | 条件, …, 走査順 | **子孫**から最初の1件 |
| `faChildren([c], [Retry], [Sleep])` | 条件（省略で全件）, … | **子**を全件（配列で保持） |
| `faDescendants(c, [Retry], [Sleep], [direction])` | 条件, …, 走査順 | **子孫**を全件 |

- `RetryCount` / `RetrySleepTime`：見つからないとき指定回数リトライ（描画待ちに有効）
- `direction`（`EnumFindDirection`）：`Default` / `PostOrder` / `LastToFirst`

```vb
' 子孫から最大5回、300ms間隔でリトライしながら探す
Set el = parent.ffDescendants(c.Type_(Button).Name4_Full("OK"), RetryCount:=5)

' 子を全件、順に名前を出す
Dim arr As uia_e, i As Long
Set arr = parent.faChildren(c.Type_(ListItem))
For i = 0 To arr.Array_Length - 1
    Debug.Print arr.Array_GetItemByIndex(i).prName
Next
```

### TreeWalker（木をたどる）

| メソッド | 説明 |
|---|---|
| `tw_Set(twType, [cnd])` | 使う walker を設定（`ControlView` / `ContentView` / `RawView` / `ByCondition`） |
| `twParent()` | 親へ |
| `twFirstChild()` | 最初の子へ |
| `twLastChild()` | 最後の子へ |
| `twNext()` | 次の兄弟へ |
| `twPrev()` | 前の兄弟へ |

`tw_Set` を呼ばずに `tw*` を使うと `ControlView` が既定で使われます。

### プロパティ（`pr*`）

| プロパティ | 型 | 説明 |
|---|---|---|
| `prName` | String | 名前 |
| `prClsName` | String | クラス名 |
| `prCtrlType` | Long | コントロール種別 ID |
| `prShortCtrlTypeName` | String | 種別の短い名前（`Button` 等） |
| `prLocalCtrlType` | String | ローカライズ種別名（`ボタン` 等） |
| `prValue` | String | ValuePattern の値 |
| `prHwnd` | LongPtr | ウィンドウハンドル |
| `prToggleState` | Long | トグル状態 |
| `prExpandCollapseState` | Long | 展開/折りたたみ状態 |
| `prScrollState_X` / `_Y` | Double | スクロール位置（%） |
| `prRect` | `tagRECT` | 矩形 **{Left, Top, Right, Bottom}** |
| `prRectLeft` / `Top` / `Right` / `Bottom` | Long | 矩形の各辺 |
| `prRectCenter` | `tagPOINT` | 矩形の中心座標 |
| `prPtnAv_Invoke` | Boolean | InvokePattern が使えるか |

```vb
Debug.Print el.prName & " [" & el.prLocalCtrlType & "] {" & el.prClsName & "}"
Debug.Print "中心: " & el.prRectCenter.X & "," & el.prRectCenter.Y
```

### 操作（`pt*` ほか）

| メソッド | パターン | 説明 |
|---|---|---|
| `ptInvoke()` | Invoke | クリック相当 |
| `ptSetValue(value)` | Value | 値を入れる |
| `ptToggle()` | Toggle | トグル |
| `ptExpand()` / `ptCollapse()` | ExpandCollapse | 展開 / 折りたたみ |
| `ptScroll([xPct], [yPct])` | Scroll | %指定でスクロール（省略で現在値維持） |
| `ptScrollIntoView()` | ScrollItem | 表示位置までスクロール |
| `ptSelectionItem_Select()` | SelectionItem | 選択 |
| `ptSelection_GetSelection()` | Selection | 選択項目を取得（`uia_e`） |
| `ptGridGetItem(row, col)` | Grid | グリッドのセルを取得（`uia_e`） |
| `ptWindowClose()` | Window | ウィンドウを閉じる |
| `ptGetTextRange()` | Text 系 | テキスト範囲（`uia_t`）を取得 |
| `SetFocus()` | ― | フォーカスを当てる |
| `ShowContextmenu()` | ― | コンテキストメニューを出す |
| `CompareTo(target)` | ― | 同じ要素か比較（Boolean） |

### 配列（`fa*` の結果を扱う）

| メソッド | 説明 |
|---|---|
| `Array_Length` | 保持している要素配列の件数 |
| `Array_GetItemByIndex(i)` | i 番目の要素（`uia_e`） |
| `ElemArray_Col` | 全要素を `Collection` で取得 |
| `Child_SetElemArray([c])` | 子要素を条件で取り込む |
| `Child_GetItem(c)` | 各子の配下を条件で探し、最初に見つかった子を返す |

### Edge 操作

現在の要素が Edge のトップウィンドウ（`Chrome_WidgetWin_1` / 名前が `*Edge`）である前提のものが多いです。`AdditionalCondition` は文字列（名前の部分一致）か `uia_c` を渡せます。

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

```vb
Dim doc As Object
Set doc = edgeTab.IES_GetIHTMLDocument_FromTopWindow()
Debug.Print doc.title
doc.getElementById("foo").Click
```

### その他 / 応用

| メソッド | 説明 |
|---|---|
| `GetTopWindow()` | 祖先をたどってトップウィンドウ（デスクトップの1つ下）を取得 |
| `ptValue_DLFile([folder], [savedPath], [overwrite])` | ValuePattern の URL を `URLDownloadToFile` で保存 |
| `GetCtrlTypeName(id)` | 種別 ID → 定数名の文字列 |
| `GetInfo([verbose])` | 要素のプロパティ/パターン一覧をイミディエイトに表示（`True` で全部） |
| `test()` | カーソル直下の要素を CTRL で拾って情報表示し続ける（ALT で終了） |
| `Sleep_(ms)` | 指定ミリ秒待つ |

---

## uia_t ― テキスト範囲

`uia_e.ptGetTextRange()` で取得します。`Returntype`（`EnumtRngSetType`）は `AsSelfInstance`（自身を書き換え）/ `AsnewInstance`（新規）。

| メソッド | 引数 | 説明 |
|---|---|---|
| `GetText([maxLength])` | 最大長（-1 で全文） | 範囲のテキストを取得 |
| `Find(returnType, text, [fromLast], [ignoreCase])` | 検索文字列ほか | テキストを検索して範囲を移動 |
| `Expand(returnType, unit)` | 単位（`EnumTextUnitSize`） | 単位境界まで範囲を広げる |
| `Move(returnType, unit, count)` | 単位, 移動量 | 範囲を移動 |
| `ScrollIntoView(alignToTop)` | Boolean | 範囲を表示位置へ |
| `Select_()` | ― | 範囲を選択 |

`EnumTextUnitSize`：`a_Character` / `b_Format` / `c_Word` / `d_Line` / `e_Paragraph` / `f_Page` / `g_Document`

---

## M_LIB_IES ― IE モードの HTML DOM（下回り）

通常は `uia_e.IES_*` 経由で使いますが、直接も呼べます。要素は生ポインタ、戻りは `Object`（`IHTMLDocument2`）。

| 関数 | 引数 | 説明 |
|---|---|---|
| `GetIES_From_NomalElement(elemPtr)` | 要素ポインタ | 配下の IES から取得 |
| `GetIES_From_ieElement(elemPtr)` | IES 要素ポインタ | IES 要素そのものから取得 |
| `GetIES(elemPtr)` | 要素ポインタ | 上の別名 |
| `GetHTMLDocumentFromIES(hwnd)` | ウィンドウハンドル | HWND から直接取得（`WM_HTML_GETOBJECT`） |

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

`ref/` 版には `UIA_ElementFromPoint.bas` という補助モジュールが含まれます。**「座標から要素を取得する」処理だけ `DispCallFunc` で書かれている**のがポイントです。

UI Automation の `ElementFromPoint` は `POINT` 構造体を**値渡し**するのですが、これが 32bit / 64bit で ABI（引数の渡り方）が異なり、そのまま COM 経由で呼ぶと 64bit でメモリ破壊を起こして Excel がクラッシュする、という厄介な箇所です。そこで、この1メソッドだけ `DispCallFunc` で直接叩いて回避しています。

| 関数 | 戻り | 説明 |
|---|---|---|
| `ElementFromPoint(pt)` | `IUIAutomationElement` | 座標（`PointAPI` 構造体）から要素 |
| `ElementFromCursor()` | `IUIAutomationElement` | マウスカーソル直下の要素 |
| `ElementFromPoint動作サンプル()` | ― | カーソルを追いながら要素名を出す動作サンプル |

**`noref/` 版では、この処理は `uia_core` に統合済み**です（`uia_core.ElementFromPoint` として実装、`POINT` の値渡しも算術で組み立てて `CopyMemory` を使わない形）。利用者は `uia_e.GetFromPoint(x, y)` / `uia_e.GetFromCursor()` から使えばよく、`UIA_ElementFromPoint.bas` を別途インポートする必要はありません。まとめると：

- **`ref/`** … 座標取得は `UIA_ElementFromPoint.bas`（別モジュール）が担当
- **`noref/`** … `uia_core` に統合。`GetFromPoint` / `GetFromCursor` から利用

---

## 補足

- `ref/`（参照あり）と `noref/`（参照なし）で API は同じですが、`noref/` は内部で `elem` / `cnd` / `tRng` を生ポインタ（`LongPtr`）として持ち、参照カウントを各クラスが管理しています。通常は意識せずに使えます。
- 詳しい導入手順とサンプルは [README](https://github.com/tarboh/uia_rap) を参照してください。
