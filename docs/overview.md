# Overview

## 目的（Why）

**EasyOpenXml.Excel** は、Open XML SDK を直接扱うことなく、
業務アプリケーションから **安全・簡潔・一貫した方法で Excel ファイルを操作する** ための
ラッパーライブラリです。

特に以下の課題を解決することを目的としています。

- Open XML SDK の API が低レベルで学習コストが高い
- Cell / Row / Worksheet / Style などの責務が分散していて見通しが悪い
- 実装者ごとに Excel 操作コードの書き方がバラつきやすい
- 業務コードに Open XML 型が露出し、保守性が低下する

---

## 解決方針（How）

本ライブラリでは、以下の設計方針を採用しています。

### 1. Open XML SDK を外部に一切露出しない

- `SpreadsheetDocument` や `WorksheetPart` などの型は **public API に出さない**
- 利用者が意識するのは `ExcelDocument` と `Pos` のみ

👉 Open XML SDK の知識がなくても Excel 操作が可能

---

### 2. 操作単位を「座標（Pos）」に集約

- セル・範囲・結合といった操作を **Pos クラス** に集約
- 値設定・書式設定・コピー／貼り付けを同一の操作モデルで提供

```csharp
excel.Pos(1, 1).Value = 123;
excel.Pos(1, 1, 3, 3).Merge();
```

👉 Excel 操作を「位置に対する操作」として統一

---

### 3. public / internal の責務分離を厳密に

- **public**：業務アプリが触る API（安定性重視）
- **internal**：Open XML SDK 直接操作（変更可能）

これにより、

- 内部実装の差し替え
- Open XML SDK のバージョン変更

を **利用者コードに影響させずに** 行える構造としています。

---

### 4. .NET Framework 4.8 を第一級ターゲットとする

- 業務システムで依然として多い **.NET Framework 4.8** を前提
- 言語機能・依存ライブラリは net48 で安定して動作するものを選択

👉 既存業務アプリへの導入を最優先

---

## 非ゴール（やらないこと）

以下は本ライブラリのスコープ外です。

- Excel の全機能を網羅すること
- 数式エンジンやグラフ生成の高度な抽象化
- UI 操作（Excel アプリケーション自動化）
- パフォーマンス最優先の大量データ処理専用ライブラリ

👉 **業務帳票生成・編集** にフォーカス

---

## 想定ユースケース

- ASP.NET / WinForms / WPF アプリからの Excel 帳票生成
- バッチ処理による Excel ファイル出力
- Excel Creator 等の代替実装

---

## 設計上のキーワード

- Facade（ExcelDocument）
- Proxy（PosProxy / PosAttrProxy）
- 情報隠蔽（Open XML SDK）
- 安定 API / 可変 internal

---



