# EasyOpenXml.Excel

.NET Framework 4.8 互換を前提とした **Open XML SDK ラッパーライブラリ** です。

本ドキュメント群は、設計意図・公開 API の使い方・内部構成・拡張方針を共有することを目的としています。

---

## 📚 ドキュメント構成

- **index.md**（このファイル）  
  ドキュメント全体の案内

- **overview.md**  
  ライブラリ概要・設計方針・非ゴール

- **architecture.md**  
  クラス構成・責務分離・レイヤ構造

- **public-api.md**  
  公開 API（ExcelDocument / Pos / PosAttr など）の仕様

- **internals.md**  
  internal クラスの役割と Open XML SDK との関係

- **lifecycle.md**  
  ファイル生成〜破棄までのライフサイクル

- **copy-paste.md**  
  Copy / Paste / Merge の内部設計

- **exceptions.md**  
  例外設計ポリシー

- **net48-notes.md**  
  .NET Framework 4.8 向け注意点

- **future.md**  
  将来拡張のアイデア（破壊的変更を伴わないもの）

---

## 🎯 想定読者

- ライブラリ利用者（業務アプリ開発者）
- 保守・改修を行うエンジニア
- 設計レビュー担当者

---

## 🧭 設計思想（要約）

- Open XML SDK を **外部に一切露出しない**
- 利用者は **ExcelDocument と Pos だけ** を理解すればよい
- 書き込み操作は「座標（Pos）」中心
- internal 層で責務を厳密に分離

---


