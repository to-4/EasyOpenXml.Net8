# Architecture

本章では **クラス構成（レイヤ）** と **責務分離**、
および **public API と internal 実装の境界** を説明します。

---

## 全体像（レイヤ構造）

本ライブラリは大きく 2 層に分かれます。

1. **Public API 層**：利用者が触る安定 API
2. **Internals 層**：Open XML SDK を直接操作する実装

```text
[利用者コード]
   |
   v
Public API
  - ExcelDocument
  - Pos / PosAttr
  - CellWrapper
  - ExcelDocumentException
   |
   v
Internals
  - ExcelInternal
  - SheetManager
  - PosProxy / PosAttrProxy / CellWrapperProxy
  - StyleManager / SharedStringManager
  - AddressConverter / Guards
  - (Copy/Paste) ExcelInternalAccessor / CellSnapshot
   |
   v
Open XML SDK
  - SpreadsheetDocument / WorkbookPart / WorksheetPart / Cell / Stylesheet ...
```

---

## Public API 層（利用者向け）

### ExcelDocument（Facade）

`ExcelDocument` はライブラリの **唯一の入口** です。

- ファイルの初期化（テンプレートコピー等）
- ブックの Open / Close / Dispose
- シート選択（名前 / インデックス）
- 値の読み書き（座標 / A1）
- `Pos(...)` による操作ハンドルの取得

**ポイント**

- 利用者は Open XML SDK の型を知らなくてよい
- 例外は `ExcelDocumentException` に集約する

---

### Pos（操作ハンドル）

`Pos` はセル／範囲を表す **操作ハンドル** です。

- 1セル指定：`Pos(r, c)`
- 範囲指定：`Pos(r1, c1, r2, c2)`

主な操作

- 値の設定：`Value`
- 文字列として書く：`Str(...)`
- 書式設定：`Attr`
- コピー：`Copy()`
- 貼り付け：`Paste()` / `Paste(r, c)`
- 結合：`Merge()`

**ポイント**

- Excel操作は「座標に対して何をするか」に集約される
- 実処理は `PosProxy` に委譲する

---

### PosAttr（書式 API）

`PosAttr` はセル／範囲のスタイル（見た目）を操作する API です。

- `FontColor`
- `BackColor`
- `Format`

**ポイント**

- Style の生成・管理は internal に隠蔽する
- 同一スタイルはキャッシュして再利用する

---

### CellWrapper（A1 指定）

`CellWrapper` は `"A1"` のような参照を扱うための薄い API です。

- `Cell("B3").Value = ...`

**ポイント**

- 内部では `AddressConverter` を使い、最終的に `PosProxy` の操作に寄せる

---

## Internals 層（実装）

Internals 層は、Open XML SDK を直接触る責務を持ちます。

### ExcelInternal（Open XML 操作の中心）

- `SpreadsheetDocument` の open/close
- Workbook / WorksheetPart の管理
- `PosProxy` 等へ必要な依存を渡す

責務の境界

- Workbook/Sheet を跨ぐ判断は `ExcelInternal`
- シートの解決ロジックは `SheetManager`

---

### SheetManager（シート解決）

- ブックからシート一覧を取得
- シート名／インデックス指定から `WorksheetPart` を解決
- 現在シート（Current）を管理

---

### PosProxy（実処理の中核）

`PosProxy` は、セルの生成・取得・値設定・スタイル反映など
**実際の Excel 変更の中心** です。

主な責務

- Row / Cell の生成（順序を壊さない）
- 値の書き込み／読み取り
- 範囲正規化（r1<=r2, c1<=c2）
- スタイルの適用（styleIndex の反映）
- セル結合（Merge）
- Copy/Paste 用のスナップショット作成

設計理由

- Open XML SDK の操作が最も複雑な部分を 1 箇所に集約する
- public API をシンプルに保つ

---

### PosAttrProxy（スタイル反映の橋渡し）

- `PosAttr` の要求を受け取る
- `StyleManager` に styleIndex を作らせる
- `PosProxy.ApplyStyle(...)` でセルに反映

---

### StyleManager（Stylesheet 管理）

- Stylesheet（Fonts/Fills/CellFormats）の生成
- 色やフォーマットに対応する styleIndex の作成
- styleIndex のキャッシュ（同一スタイルの再利用）

注意

- 色の管理は `ARGB` をキーに統一するとよい

---

### SharedStringManager（SharedStringTable 管理）

- 文字列を書き込む際の SharedStringTable を管理
- 既存文字列の検索 / 追加

改善余地

- 線形探索になりやすいので内部キャッシュを検討

---

### AddressConverter（A1 変換）

- 行列（r,c）⇔ A1（"B3"）の相互変換

---

### Copy/Paste：ExcelInternalAccessor / CellSnapshot

Copy/Paste は「クリップボード」を internal で保持します。

- `Copy()`：対象範囲を `CellSnapshot` として保持
- `Paste()`：保持した Snapshot を使って貼り付け

設計ポイント

- Copy/Paste の状態を public API に出さない
- ExcelDocument インスタンス（内部の SpreadsheetDocument）に紐付ける

注意

- 破棄（Dispose）時に Snapshot も掃除する（リーク対策）

---

## public / internal 境界（ポリシー）

### 原則

- public API は **安定** であること
- internal は **自由に変更可能** であること
- Open XML SDK の型は public に出さない

### 例外設計

- Open XML SDK 由来の例外は public に漏らさない
- `ExcelDocumentException` に包んで利用者へ返す

---

## 典型フロー（Initialize → 操作 → Finalize）

```text
ExcelDocument.InitializeFile
  -> ExcelInternal.Open
  -> SheetManager.Load

ExcelDocument.Pos(...)
  -> new Pos(PosProxy)

Pos.Value/Attr/Merge/Copy/Paste
  -> PosProxy / PosAttrProxy
  -> SharedStringManager / StyleManager

ExcelDocument.FinalizeFile
  -> ExcelInternal.Save
  -> ExcelInternal.Dispose
```

---

## .NET Framework 4.8 への適用メモ

- `init` アクセサ等の新しめ言語機能は避け、`private set` / ctor で表現する
- implicit usings に依存せず、必要な `using` を明示する
- `System.Drawing.Common` ではなく `System.Drawing` を使用する

---

次に読む → [public-api.md]

