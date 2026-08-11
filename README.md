# uia_rap

UI Automation を使って **Edge などのアプリを VBA から自動操作する** ための、チェーン可能で実用重視のラッパライブラリ。

作者（[tarboh](https://github.com/tarboh)）が普段使いしている機能に絞ってあり、UI Automation の全機能を網羅するものではありません。「よく使う操作を短く書く」ことを目的にしています。

- `uia_e` … 要素（Element）。取得・検索・ツリー走査・プロパティ・パターン操作
- `uia_c` … 検索条件（Condition）。名前・種類・クラス名などをチェーンで組み立てる
- `uia_t` … テキスト範囲（TextRange）。検索・移動・選択
- `uia_Factory` … `e()` / `c()` / `t()` で各インスタンスを生成する入口

```vb
Sub Example()
    Dim el As uia_e
    ' 電卓の「7」ボタンを名前と種類で探して押す
    Set el = e.getRoot.ffDescendants(c.Type_(Button).Name4_Full("7"))
    el.ptInvoke
End Sub
```

---

## 導入

このバージョン（`ref/`）は **参照設定が前提** です。VBE の「ツール → 参照設定」で次の2つにチェックを入れてください。

| 参照設定 | 用途 |
|---|---|
| **UIAutomationClient** | UI Automation 本体（必須） |
| **Microsoft HTML Object Library** | `M_LIB_IES` / Edge・IE の HTML DOM 抽出（`IHTMLDocument2`）で使用 |

そのうえで `ref/` の6ファイルをインポートします。

| ファイル | 役割 |
|---|---|
| `uia_e.cls` | 要素クラス |
| `uia_c.cls` | 条件クラス |
| `uia_t.cls` | テキスト範囲クラス |
| `uia_Factory.bas` | `e()` / `c()` / `t()` ファクトリ |
| `UIA_ElementFromPoint.bas` | 座標から要素を取る補助（`DispCallFunc` 版） |
| `M_LIB_IES.bas` | IE/旧 Edge の HTML DOM 抽出 |

> **参照設定不要版を作成中です。** 詳細は下の «参照設定不要版» を参照。

---

## 使い方

### 要素を取得する

```vb
Set el = e.getRoot                 ' デスクトップ（ルート）
Set el = e.getFocus                ' フォーカス中の要素
Set el = e.getHandle(hwnd)         ' ウィンドウハンドルから
Set el = e.GetFromCursor           ' マウスカーソル直下
Set el = e.GetFromPoint(x, y)      ' 座標から
```

### 条件を組み立てる（`uia_c`）

条件はチェーンで積み上げます。既定は AND 結合。

```vb
' 種類=Button かつ 名前="保存"（完全一致）
Set cond = c.Type_(Button).Name4_Full("保存")

' 名前に "OK" を含む（部分一致・大小無視）
Set cond = c.Name1_Sub("OK")

' 名前が "はい" または "OK"
Set cond = c.Name2_Sub_Or("はい", "OK")

' クラス名 / AutomationId で
Set cond = c.ClsName("Button").AutomationId("saveBtn")
```

`Style` 引数で `AsAnd` / `AsOr` / `AsNew` を切り替えられます。

### 検索する（`uia_e`）

```vb
' 子から最初の1件
Set el = parent.ffChildren(cond)

' 子孫から最初の1件（見つかるまで最大5回リトライ、間隔300ms）
Set el = parent.ffDescendants(cond, RetryCount:=5, RetrySleepTime:=300)

' 子孫を全部（配列として保持）
Set els = parent.faDescendants(cond)
Dim i As Long
For i = 0 To els.Array_Length - 1
    Debug.Print els.Array_GetItemByIndex(i).prName
Next i
```

### ツリーを歩く（TreeWalker）

```vb
el.tw_Set ControlView              ' ContentView / RawView / ByCondition も可
Set parent = el.twParent
Set first  = el.twFirstChild
Set next_  = el.twNext
```

### プロパティを読む

```vb
Debug.Print el.prName              ' 名前
Debug.Print el.prClsName           ' クラス名
Debug.Print el.prCtrlType          ' コントロール種別 (ID)
Debug.Print el.prLocalCtrlType     ' ローカライズされた種別名
Debug.Print el.prValue             ' ValuePattern の値
Debug.Print el.prToggleState       ' トグル状態
Debug.Print el.prHwnd              ' ウィンドウハンドル
Dim r As tagRECT: r = el.prRect    ' 矩形
Debug.Print el.prRectCenter.x, el.prRectCenter.y   ' 中心座標
```

### 操作する（パターン）

```vb
el.ptInvoke                        ' クリック相当（InvokePattern）
el.ptSetValue "abc"                ' 値を入れる（ValuePattern）
el.ptToggle                        ' トグル
el.ptExpand : el.ptCollapse        ' 展開・折りたたみ
el.ptScroll 50, 50                 ' スクロール（%指定）
el.ptScrollIntoView                ' 表示位置までスクロール
el.ptSelectionItem_Select          ' 選択
el.ptWindowClose                   ' ウィンドウを閉じる
el.SetFocus                        ' フォーカスを当てる
```

---

## メモ

- チェーンの起点 `e` / `c` / `t` は `uia_Factory` の関数で、呼ぶたびに新しいインスタンスを返します。
- `uia_c.cnd` は既定メンバー（`VB_VarUserMemId = 0`）で、生の `IUIAutomationCondition` を保持します。
- Edge 操作用の内部ヘルパ（`EdgeGetTopWindow` など）はクラス内の Private 関数として持っています。
- `M_LIB_IES` は IE / 旧 Edge（EdgeHTML）向けの HTML DOM 抽出です。**現在の Chromium 版 Edge ではこの経路は使えません**。

---

## 参照設定不要版（作成中）

`ref/` は参照設定が前提ですが、**参照設定なし版**（`noref/`）を作成中です。設計方針は次のとおり。

- `CoCreateInstance` + `DispCallFunc` で UI Automation を叩き、**参照設定を不要**にする。
- インスタンス生成機構（`New CUIAutomation` 相当）は `uia_e` などのクラス内に隠蔽する。
- 「インポートするモジュールが少ない」利点を保つため、必要な呼び出しだけをクラスに内蔵し、**モジュール数を増やさない**。
- MSHTML（`IHTMLDocument2`）部分は遅延バインドで残す。

同じ姉妹プロジェクトとして、UI Automation を参照設定なしでフルに叩くライブラリもあります → **[VBA_UIAutomation_NoRef](https://github.com/tarboh/VBA_UIAutomation_NoRef)**

---

## ライセンス

MIT
