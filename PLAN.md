# Open XML SDK TypeScript 完全互換実装 計画書

## 概要

本計画書は、Microsoft の [Open XML SDK](https://github.com/dotnet/Open-XML-SDK)（C#/.NET 実装）と完全互換となる TypeScript ライブラリの設計・実装方針を定めるものです。OOXML（Office Open XML）規格 [ECMA-376](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/) に準拠し、DOCX・XLSX・PPTX ファイルの生成・読み取り・編集を TypeScript/JavaScript 環境で可能にします。

---

## 1. 目標と対象範囲

### 1.1 目標
- Open XML SDK の公開 API を TypeScript で完全再現する
- Node.js および最新ブラウザ（ESM）の両環境で動作する
- 型安全性を最大限に活かした強型付き要素クラス群を提供する
- ゼロ依存（コア機能）または最小依存でランタイムを軽量に保つ
- 既存の C# コードをほぼ 1:1 で TypeScript へ移植できるAPIを提供する

### 1.2 対象ファイル形式
| 形式 | ドキュメント種別 | 名前空間プレフィックス |
|------|-----------------|----------------------|
| DOCX / DOTX | Word 文書 | `w:` (WordprocessingML) |
| XLSX / XLTX | Excel スプレッドシート | `x:` (SpreadsheetML) |
| PPTX / POTX | PowerPoint プレゼンテーション | `p:` (PresentationML) |
| VSDX | Visio 図面 | `v:` (DrawingML 拡張) |
| 共通 | 図形・グラフ・画像 | DrawingML, Shared ML |

### 1.3 対象外（将来バージョン）
- マクロ付きファイル（XLSM, DOCM など）のマクロ実行
- PDF エクスポート
- OLE 埋め込みオブジェクトの操作

---

## 2. アーキテクチャ設計

### 2.1 パッケージ構成

```
@openxmlsdk/
├── core          # パッケージ・パート・関係・コンテンツタイプ基盤
├── wordprocessing # WordprocessingML (DOCX)
├── spreadsheet   # SpreadsheetML (XLSX)
├── presentation  # PresentationML (PPTX)
├── drawing       # DrawingML 共通要素
├── shared-types  # OOXML 共有型定義 (ECMA-376 Part 1 & 4)
└── validator     # スキーマ検証エンジン
```

### 2.2 レイヤー構造

```
┌─────────────────────────────────────────────────────┐
│               アプリケーション層                      │
│   (WordprocessingDocument / SpreadsheetDocument …)  │
├─────────────────────────────────────────────────────┤
│               ドメイン要素層                          │
│   (Paragraph / Table / Cell / Shape …)              │
├─────────────────────────────────────────────────────┤
│               OpenXmlElement 基底層                  │
│   (OpenXmlElement / OpenXmlCompositeElement …)      │
├─────────────────────────────────────────────────────┤
│               パッケージ層                           │
│   (OpenXmlPackage / OpenXmlPart / Relationships …)  │
├─────────────────────────────────────────────────────┤
│               ZIP / XML I/O 層                       │
│   (fflate / fast-xml-parser またはネイティブ API)     │
└─────────────────────────────────────────────────────┘
```

### 2.3 コアクラス階層

```
OpenXmlElement
├── OpenXmlLeafElement          # 子要素を持たない葉ノード
├── OpenXmlLeafTextElement      # テキスト値を持つ葉ノード
└── OpenXmlCompositeElement     # 子要素を持つ複合ノード
    ├── OpenXmlPartRootElement  # パートのルート要素
    └── (各ドメイン要素クラス)

OpenXmlPackage                  # ZIP コンテナ抽象
├── WordprocessingDocument
├── SpreadsheetDocument
└── PresentationDocument

OpenXmlPart                     # パート（XML ファイル）抽象
├── MainDocumentPart
├── WorkbookPart
└── PresentationPart
    └── (各サブパート)
```

---

## 3. モジュール詳細設計

### 3.1 `@openxmlsdk/core`

#### 3.1.1 ZIP パッケージ管理

```typescript
interface IPackageStorage {
  readPart(uri: string): Promise<Uint8Array>;
  writePart(uri: string, data: Uint8Array): Promise<void>;
  deletePart(uri: string): Promise<void>;
  listParts(): Promise<string[]>;
  flush(): Promise<Uint8Array>; // バイト列として返す
}

class InMemoryPackageStorage implements IPackageStorage { ... }
class StreamPackageStorage implements IPackageStorage { ... }  // Node.js専用
```

- ZIP 操作: `fflate`（ブラウザ/Node.js 共通、ESM対応）を採用
- Node.js 環境では `fs` を使ったストリーミング読み書きも対応

#### 3.1.2 コンテンツタイプ・関係管理

```typescript
class ContentTypeCollection {
  addDefault(extension: string, contentType: string): void;
  addOverride(partName: string, contentType: string): void;
  getContentType(partName: string): string | undefined;
}

class PackageRelationshipCollection {
  addRelationship(opts: RelationshipOptions): PackageRelationship;
  removeRelationship(id: string): void;
  getById(id: string): PackageRelationship | undefined;
  getByType(type: string): PackageRelationship[];
}
```

#### 3.1.3 OpenXmlElement 基底

```typescript
abstract class OpenXmlElement {
  // C# SDK 互換プロパティ
  parent: OpenXmlElement | undefined;
  localName: string;
  namespaceUri: string;
  prefix: string;
  innerText: string;
  outerXml: string;

  // 子要素操作
  append(...children: OpenXmlElement[]): this;
  prepend(...children: OpenXmlElement[]): this;
  insertBefore<T extends OpenXmlElement>(newChild: T, refChild: OpenXmlElement | undefined): T;
  insertAfter<T extends OpenXmlElement>(newChild: T, refChild: OpenXmlElement | undefined): T;
  removeChild<T extends OpenXmlElement>(child: T): T;
  removeAllChildren(): void;

  // 走査
  elements<T extends OpenXmlElement>(type?: Constructor<T>): Iterable<T>;
  descendants<T extends OpenXmlElement>(type?: Constructor<T>): Iterable<T>;
  ancestors<T extends OpenXmlElement>(type?: Constructor<T>): Iterable<T>;

  // クローン
  clone(deep?: boolean): this;

  // シリアライズ
  toXmlString(options?: SerializeOptions): string;
  static fromXmlString<T extends OpenXmlElement>(xml: string): T;
}
```

#### 3.1.4 属性システム

```typescript
// C# SDK の StringValue / IntegerValue 等に相当する型ラッパー
class OpenXmlSimpleType<T> {
  value: T | undefined;
  hasValue: boolean;
  toString(): string;
  static fromString<T>(str: string): OpenXmlSimpleType<T>;
}

type StringValue  = OpenXmlSimpleType<string>;
type IntegerValue = OpenXmlSimpleType<number>;
type BooleanValue = OpenXmlSimpleType<boolean>;
type EnumValue<T extends string> = OpenXmlSimpleType<T>;
```

### 3.2 `@openxmlsdk/wordprocessing`

主要クラス（C# SDK 互換）:

| TypeScript クラス | 対応 XML 要素 | 説明 |
|------------------|--------------|------|
| `Body` | `w:body` | ドキュメント本文 |
| `Paragraph` | `w:p` | 段落 |
| `Run` | `w:r` | テキストラン |
| `Text` | `w:t` | テキスト内容 |
| `Table` | `w:tbl` | 表 |
| `TableRow` | `w:tr` | 表の行 |
| `TableCell` | `w:tc` | 表のセル |
| `ParagraphProperties` | `w:pPr` | 段落プロパティ |
| `RunProperties` | `w:rPr` | ランプロパティ |
| `BookmarkStart` / `BookmarkEnd` | `w:bookmarkStart` / `End` | ブックマーク |
| `SectionProperties` | `w:sectPr` | セクションプロパティ |
| `Style` | `w:style` | スタイル定義 |
| `Numbering` | `w:numbering` | リスト番号付け |
| `AbstractNum` | `w:abstractNum` | 抽象番号付け |

### 3.3 `@openxmlsdk/spreadsheet`

主要クラス:

| TypeScript クラス | 対応 XML 要素 | 説明 |
|------------------|--------------|------|
| `Workbook` | `x:workbook` | ワークブック |
| `Sheet` | `x:sheet` | シート参照 |
| `Worksheet` | `x:worksheet` | ワークシート |
| `SheetData` | `x:sheetData` | データ領域 |
| `Row` | `x:row` | 行 |
| `Cell` | `x:c` | セル |
| `CellValue` | `x:v` | セル値 |
| `SharedStringTable` | `x:sst` | 共有文字列テーブル |
| `StyleSheet` | `x:styleSheet` | スタイルシート |
| `MergeCells` / `MergeCell` | `x:mergeCells` / `x:mergeCell` | セル結合 |
| `ConditionalFormatting` | `x:conditionalFormatting` | 条件付き書式 |
| `DefinedName` | `x:definedName` | 名前定義（数式等） |

### 3.4 `@openxmlsdk/presentation`

主要クラス:

| TypeScript クラス | 対応 XML 要素 | 説明 |
|------------------|--------------|------|
| `Presentation` | `p:presentation` | プレゼンテーション |
| `SlideMasterIdList` | `p:sldMasterIdLst` | スライドマスター一覧 |
| `SlideIdList` | `p:sldIdLst` | スライド一覧 |
| `Slide` | `p:sld` | スライド |
| `SlideLayout` | `p:sldLayout` | スライドレイアウト |
| `SlideMaster` | `p:sldMaster` | スライドマスター |
| `CommonSlideData` | `p:cSld` | スライド共通データ |
| `ShapeTree` | `p:spTree` | 図形ツリー |
| `Shape` | `p:sp` | 図形 |
| `Picture` | `p:pic` | 画像 |

### 3.5 `@openxmlsdk/drawing`

DrawingML 共通要素（Word/Excel/PowerPoint で共有）:

- `NonVisualDrawingProperties` / `NonVisualPictureProperties`
- `BlipFill` / `Blip`
- `Transform2D` / `Offset` / `Extents`
- `SolidFill` / `GradientFill` / `PatternFill`
- `Outline` / `EffectList`
- `Chart` 系（`ChartReference` など）

---

## 4. 実装フェーズ

### Phase 1: コア基盤（優先度: 最高）

**期間目安: 6〜8週間**

- [ ] プロジェクト初期設定（TypeScript 5.x, ESM, monorepo with pnpm workspaces）
- [ ] `@openxmlsdk/core` パッケージ
  - [ ] ZIP I/O 層（fflate ベース）
  - [ ] `OpenXmlElement` 基底クラス
  - [ ] 属性型システム (`StringValue`, `IntegerValue` 等)
  - [ ] XML シリアライズ・デシリアライズ
  - [ ] `OpenXmlPackage` 抽象クラス
  - [ ] `OpenXmlPart` 抽象クラス
  - [ ] コンテンツタイプ管理
  - [ ] 関係管理 (`_rels` ディレクトリ)
- [ ] 単体テスト（Vitest）: カバレッジ 90% 以上

### Phase 2: WordprocessingML（優先度: 高）

**期間目安: 8〜10週間**

- [ ] `WordprocessingDocument` クラス
- [ ] `MainDocumentPart` および関連パート
  - [ ] `StyleDefinitionsPart`
  - [ ] `NumberingDefinitionsPart`
  - [ ] `FontTablePart`
  - [ ] `HeaderPart` / `FooterPart`
  - [ ] `CommentsPart`
  - [ ] `FootnotesPart` / `EndnotesPart`
  - [ ] `ImagePart`（埋め込み画像）
- [ ] すべての `w:` 要素クラスの自動コード生成（ECMA-376 XSD から）
- [ ] DOCX 読み取り・書き込みの結合テスト
- [ ] 実際の Word ファイルを使ったリグレッションテスト

### Phase 3: SpreadsheetML（優先度: 高）

**期間目安: 8〜10週間**

- [ ] `SpreadsheetDocument` クラス
- [ ] `WorkbookPart` および関連パート
  - [ ] `WorksheetPart`
  - [ ] `SharedStringTablePart`
  - [ ] `WorkbookStylesPart`
  - [ ] `ChartPart`
  - [ ] `DrawingsPart`
- [ ] セル値型の完全サポート（数値・文字列・日付・ブール・エラー）
- [ ] 数式文字列の保存・読み込み（評価は対象外）
- [ ] XLSX 読み取り・書き込みの結合テスト

### Phase 4: PresentationML（優先度: 中）

**期間目安: 6〜8週間**

- [ ] `PresentationDocument` クラス
- [ ] `PresentationPart` および関連パート
  - [ ] `SlidePart`
  - [ ] `SlideLayoutPart`
  - [ ] `SlideMasterPart`
  - [ ] `NotesSlidePart`
- [ ] PPTX 読み取り・書き込みの結合テスト

### Phase 5: 高度機能・品質向上（優先度: 中）

**期間目安: 4〜6週間**

- [ ] `@openxmlsdk/validator` スキーマ検証
- [ ] `OpenXmlPackage.validate()` メソッド
- [ ] フラグメント操作 API（`DocumentFormat.OpenXml.Packaging` 相当）
- [ ] `FlatOpc` 形式のサポート（XMLとしてシリアライズ）
- [ ] パフォーマンス最適化（大規模ファイル処理）

### Phase 6: ドキュメント・エコシステム（優先度: 低）

**期間目安: 4〜6週間**

- [ ] API ドキュメント（TypeDoc）
- [ ] 移行ガイド（C# SDK → TypeScript SDK）
- [ ] コード例集（レシピ）
- [ ] CLI ツール（`@openxmlsdk/cli`）
- [ ] npm パッケージ公開設定

---

## 5. 技術スタック

### 5.1 言語・ビルド

| 項目 | 採用技術 | 理由 |
|------|---------|------|
| 言語 | TypeScript 5.x | 強型付き・デコレータ・ESM完全対応 |
| ビルドツール | tsup / rollup | ESM + CJS デュアルビルド |
| パッケージマネージャ | pnpm + workspaces | monorepo管理・高速インストール |
| ターゲット | ES2020 + ESM | Node.js 18+, 主要ブラウザ対応 |

### 5.2 実行時依存（コア）

| ライブラリ | 用途 | バンドルサイズ |
|-----------|------|--------------|
| `fflate` | ZIP 圧縮・展開 | ~18KB gzip |
| `fast-xml-parser` | XML 解析 | ~43KB gzip |

### 5.3 開発時依存

| ライブラリ | 用途 |
|-----------|------|
| `vitest` | 単体・結合テスト |
| `typedoc` | API ドキュメント生成 |
| `eslint` + `@typescript-eslint` | 静的解析 |
| `prettier` | コードフォーマット |
| `changesets` | バージョン管理・CHANGELOG |

### 5.4 コード生成

ECMA-376 XSD スキーマから要素クラスを自動生成するコードジェネレータを実装する。

```
tools/
└── codegen/
    ├── parse-xsd.ts      # XSD パーサー
    ├── generate-class.ts # TypeScript クラス生成
    └── generate-enum.ts  # 列挙型生成
```

生成対象:
- 約 5,000 以上の OOXML 要素クラス
- 属性プロパティ（型付き getter/setter）
- 子要素のカーディナリティ定義
- XML 名前空間・ローカル名マッピング

---

## 6. API 設計方針

### 6.1 C# SDK との互換性原則

```typescript
// C# SDK のコード:
// using (var doc = WordprocessingDocument.Open("file.docx", true))
// {
//     var body = doc.MainDocumentPart.Document.Body;
//     body.AppendChild(new Paragraph(new Run(new Text("Hello"))));
// }

// TypeScript SDK の等価コード:
const doc = await WordprocessingDocument.open(buffer, true);
const body = doc.mainDocumentPart!.document!.body!;
body.append(new Paragraph(new Run(new Text("Hello"))));
await doc.save();
```

### 6.2 命名規則

| C# SDK | TypeScript SDK | 規則 |
|--------|----------------|------|
| `MainDocumentPart` | `MainDocumentPart` | パスカルケース維持 |
| `AppendChild()` | `append()` | Web標準に合わせキャメルケース |
| `InnerText` | `innerText` | キャメルケース |
| `OuterXml` | `outerXml` | キャメルケース |
| `HasChildren` | `hasChildren` | キャメルケース |
| `new Paragraph()` | `new Paragraph()` | コンストラクタは同一 |

### 6.3 非同期 API 方針

- ファイル I/O（open/save/flush）は `async/await`
- DOM 操作（要素追加・取得・変更）は同期
- ストリーミング大規模ファイル処理は `AsyncIterable` で提供

### 6.4 エラーハンドリング

```typescript
class OpenXmlException extends Error { ... }
class InvalidOperationException extends OpenXmlException { ... }
class InvalidOpenXmlDocumentException extends OpenXmlException { ... }
class ValidationException extends OpenXmlException {
  errors: ValidationError[];
}
```

---

## 7. テスト戦略

### 7.1 テスト分類

| 種別 | ツール | 説明 |
|------|--------|------|
| 単体テスト | Vitest | 各クラス・メソッドの動作検証 |
| スナップショットテスト | Vitest | XML 出力の回帰テスト |
| 結合テスト | Vitest | パッケージ全体の読み書き検証 |
| 互換性テスト | カスタム | 実際の Office ファイルとの往復テスト |
| パフォーマンステスト | vitest-bench | 大規模ファイル処理のベンチマーク |

### 7.2 互換性テスト方針

- Microsoft Office で作成した本物の DOCX/XLSX/PPTX をテスト資産として使用
- TypeScript SDK で読み取り → 変更なし → 書き出し → バイト比較（構造的等価性）
- LibreOffice / WPS Office でも開けることを確認

### 7.3 カバレッジ目標

| パッケージ | 目標カバレッジ |
|-----------|--------------|
| `core` | 95% 以上 |
| `wordprocessing` | 85% 以上 |
| `spreadsheet` | 85% 以上 |
| `presentation` | 80% 以上 |
| 全体 | 85% 以上 |

---

## 8. ディレクトリ構造（最終形）

```
openxmlsdk-typescript/
├── packages/
│   ├── core/
│   │   ├── src/
│   │   │   ├── packaging/        # ZIP・パート・関係
│   │   │   ├── elements/         # OpenXmlElement 基底
│   │   │   ├── attributes/       # 属性型システム
│   │   │   ├── serialization/    # XML シリアライズ
│   │   │   └── index.ts
│   │   ├── test/
│   │   ├── package.json
│   │   └── tsconfig.json
│   ├── wordprocessing/
│   │   ├── src/
│   │   │   ├── parts/            # WordprocessingML パート
│   │   │   ├── elements/         # w: 要素クラス（自動生成含む）
│   │   │   ├── generated/        # XSD から自動生成
│   │   │   └── index.ts
│   │   └── ...
│   ├── spreadsheet/
│   ├── presentation/
│   ├── drawing/
│   ├── shared-types/
│   └── validator/
├── tools/
│   └── codegen/                  # XSD → TypeScript コードジェネレータ
├── examples/
│   ├── hello-world-docx/
│   ├── excel-report/
│   └── presentation-builder/
├── docs/
├── pnpm-workspace.yaml
├── package.json
├── tsconfig.base.json
├── vitest.config.ts
└── PLAN.md
```

---

## 9. マイルストーン・成功指標

| マイルストーン | 成功指標 |
|--------------|---------|
| M1: コア基盤 v0.1 | DOCX の最小限の読み書きが可能 |
| M2: Word v0.5 | Open XML SDK の Word テストスイートを 80% 以上通過 |
| M3: Excel v0.5 | Open XML SDK の Excel テストスイートを 80% 以上通過 |
| M4: PowerPoint v0.5 | Open XML SDK の PPT テストスイートを 75% 以上通過 |
| M5: v1.0 GA | 全体カバレッジ 85%・ドキュメント完備・npm 公開 |

---

## 10. リスクと対策

| リスク | 影響度 | 対策 |
|--------|--------|------|
| ECMA-376 の複雑性（仕様書 6,000+ ページ） | 高 | XSD からのコード自動生成で対応 |
| ブラウザでの ZIP/XML 処理パフォーマンス | 中 | Web Workers へのオフロード・ストリーミング API |
| 大規模ファイル（100MB+）のメモリ使用 | 中 | SAX スタイルの遅延解析オプションの提供 |
| Microsoft 独自拡張（非標準属性） | 中 | 拡張属性を保持するパススルー機構の実装 |
| C# SDK の非公開内部動作への依存 | 低 | 公開 API テストとリバースエンジニアリングで対応 |

---

## 参考資料

- [ECMA-376 標準仕様](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)
- [Open XML SDK GitHub（C# 実装）](https://github.com/dotnet/Open-XML-SDK)
- [Open XML SDK ドキュメント](https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk)
- [Open XML Explained（Eric White）](https://ericwhite.com/blog/open-xml-developer/)
- [fflate](https://github.com/101arrowz/fflate)
- [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser)
