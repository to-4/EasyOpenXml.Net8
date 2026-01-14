# EasyOpenXml.Excel

C# / .NET で **Open XML SDK を簡単に扱うための学習・検証用プロジェクト**です。  
Excel（.xlsx）ファイルの読み書きを、できるだけシンプルな API で行うことを目的としています。

---

## 概要

このリポジトリは、Open XML SDK を直接使う際の

- 記述量が多い
- 学習コストが高い
- 目的の処理が分かりにくい

といった課題を解消するために、  
**ラッパー設計・サンプル実装・デモコンソール**を通して理解を深めることを目的としています。

---

## プロジェクト構成

```text
EasyOpenXml.Excel.Net8/
├─ EasyOpenXml.Excel.Net8.sln
├─ src/
│  ├─ EasyOpenXml.Excel/
│  │  └─ （ライブラリ本体）
│  └─ EasyOpenXml.Excel.DemoConsole/
│     └─ （動作確認用コンソールアプリ）
└─ README.md

## 動作環境と制限事項

本プロジェクトでは `.NET 8` を使用していますが、依存ライブラリの仕様により環境ごとに以下の制約があります。

- **Windows 環境**
  - `System.Drawing.Common` および `DocumentFormat.OpenXml` が必要です。
  - すべての機能（フォントカラーの変更等を含む）が利用可能です。
- **macOS / Linux 環境**
  - .NET 6 以降、`System.Drawing.Common` は Windows 以外でサポートされなくなったため、**フォントカラーの変更などの描画に依存する機能は実装（動作）できません。**
  - 基本的なセルの読み書きには `DocumentFormat.OpenXml` が必要です。
