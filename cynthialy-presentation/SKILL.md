---
name: cynthialy-presentation
description: Cynthialyブランドのプロフェッショナルなプレゼンテーションを作成するスキル。ダークグリーン(#0D2623)とクリームイエロー(#E8DE9F)を基調とした、統一感のあるビジネス向けスライドを生成します。pptxgenjsを使用してNode.jsで実行。8種類のスライドパターンに加え、7種類のチャートパターン（縦棒・積み上げ棒・横棒・ウォーターフォール・バブル・折れ線・スコアカード）を搭載。
metadata:
  version: 2.0.0
  language: javascript
  runtime: node
---

# Cynthialy Presentation Skill

## 概要

このスキルは、Cynthialy株式会社のブランドカラーとデザインシステムに基づいたプロフェッショナルなPowerPointプレゼンテーションを作成します。

**使用するケース**:
- ビジネス提案資料の作成
- AIエージェント導入に関するプレゼンテーション
- クライアント向けの営業資料
- 統計データを含むスライド
- 構造的な課題と解決策を説明するスライド

## 技術仕様

- **ライブラリ**: pptxgenjs (Node.js)
- **スライドサイズ**: 16:9 (LAYOUT_16x9)
- **フォント**: Noto Sans JP

## カラーパレット

```javascript
const DARK_GREEN = "0D2623";       // メイン背景、見出し
const CREAM_YELLOW = "E8DE9F";     // アクセント、強調
const LIGHT_GRAY = "F3F3F3";       // コンテンツボックス背景
const TEXT_DARK = "0D2623";        // 見出しテキスト
const TEXT_MEDIUM = "434343";      // 本文テキスト
const TEXT_LIGHT = "595959";       // 補足テキスト（出典など）
const WHITE = "FFFFFF";            // 白テキスト
const HIGHLIGHT_YELLOW = "FFFF00"; // タイトル下線
const DARKER_GREEN = "1A3D38";     // ダークグリーン背景内のボックス
```

## スライドパターン

### パターン1: タイトルスライド

ダークグリーン背景、クリームイエローのアクセントライン付き

```javascript
{
  let slide = pres.addSlide();
  slide.background = { color: DARK_GREEN };
  
  // アクセントライン
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.3, w: 2, h: 0.06,
    fill: { color: CREAM_YELLOW }
  });
  
  // メインタイトル
  slide.addText("タイトル文", {
    x: 0.5, y: 1.5, w: 9, h: 1.8,
    fontSize: 36, fontFace: "Noto Sans JP", bold: true,
    color: WHITE, align: "left", valign: "top", margin: 0, lineSpacing: 50
  });
  
  // サブタイトル
  slide.addText("サブタイトル説明文", {
    x: 0.5, y: 3.5, w: 9, h: 1.2,
    fontSize: 13, fontFace: "Noto Sans JP",
    color: WHITE, align: "left", valign: "top", margin: 0
  });
  
  // 会社名
  slide.addText("Cynthialy株式会社", {
    x: 0.5, y: 4.9, w: 5, h: 0.4,
    fontSize: 14, fontFace: "Noto Sans JP",
    color: CREAM_YELLOW, align: "left", valign: "middle", margin: 0
  });
}
```

### パターン2: 統計ボックス付きスライド

3つの統計データを並べて表示

```javascript
// ヘルパー関数
function addStatBox(slide, stat, label, x, y, w, h) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: x, y: y, w: w, h: h,
    fill: { color: LIGHT_GRAY }
  });
  slide.addText(stat, {
    x: x, y: y + 0.15, w: w, h: h * 0.5,
    fontSize: 36, fontFace: "Noto Sans JP", bold: true,
    color: DARK_GREEN, align: "center", valign: "middle", margin: 0
  });
  slide.addText(label, {
    x: x + 0.1, y: y + h * 0.55, w: w - 0.2, h: h * 0.4,
    fontSize: 11, fontFace: "Noto Sans JP",
    color: TEXT_MEDIUM, align: "center", valign: "top", margin: 0
  });
}

// 使用例
{
  let slide = pres.addSlide();
  slide.background = { color: WHITE };
  
  // タイトル
  slide.addText("スライドタイトル", {
    x: 0.5, y: 0.3, w: 9, h: 0.5,
    fontSize: 22, fontFace: "Noto Sans JP", bold: true,
    color: DARK_GREEN, align: "left", valign: "middle", margin: 0
  });
  
  // 黄色のアンダーライン
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 0.85, w: 5.5, h: 0.05,
    fill: { color: HIGHLIGHT_YELLOW }
  });
  
  // 統計ボックス（3列）
  addStatBox(slide, "88%", "グローバルで\nAIを業務利用する企業", 0.5, 2.0, 2.8, 1.3);
  addStatBox(slide, "74%", "PoCから\n価値創出に至っていない", 3.5, 2.0, 2.8, 1.3);
  addStatBox(slide, "4%", "AI投資から\n十分な価値を得ている", 6.5, 2.0, 2.8, 1.3);
  
  // 出典
  slide.addText("出典：BCG「Where's the Value in AI?」2024年10月", {
    x: 0.5, y: 5.2, w: 9, h: 0.3,
    fontSize: 9, fontFace: "Noto Sans JP",
    color: TEXT_LIGHT, align: "left", valign: "middle", margin: 0
  });
}
```

### パターン3: 番号付きリストスライド（5項目）

```javascript
// ヘルパー関数
function addBadge(slide, num, x, y, size = 0.5) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: size, h: size,
    fill: { color: DARK_GREEN }
  });
  slide.addText(String(num), {
    x: x, y: y, w: size, h: size,
    fontSize: 18, fontFace: "Noto Sans JP", bold: true,
    color: WHITE, align: "center", valign: "middle", margin: 0
  });
}

// 使用例
{
  let slide = pres.addSlide();
  slide.background = { color: WHITE };
  
  slide.addText("5つの構造的な原因", {
    x: 0.5, y: 0.3, w: 9, h: 0.5,
    fontSize: 22, fontFace: "Noto Sans JP", bold: true,
    color: DARK_GREEN, align: "left", valign: "middle", margin: 0
  });
  
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 0.85, w: 5.5, h: 0.05,
    fill: { color: HIGHLIGHT_YELLOW }
  });
  
  // クリームイエローのキーメッセージボックス
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.1, w: 9, h: 0.6,
    fill: { color: CREAM_YELLOW }
  });
  slide.addText("前提条件の欠落は、大きく次の5つに整理できる", {
    x: 0.7, y: 1.1, w: 8.6, h: 0.6,
    fontSize: 13, fontFace: "Noto Sans JP", bold: true,
    color: DARK_GREEN, align: "left", valign: "middle", margin: 0
  });
  
  // リスト項目
  const items = [
    { title: "絵姿不在", desc: "ゴールイメージがない" },
    { title: "業務・暗黙知の未整理", desc: "可視化されていない" },
    { title: "データ/ナレッジ基盤不足", desc: "整備されていない" },
    { title: "ガバナンス不備", desc: "未整備" },
    { title: "人材・体制の不在", desc: "人材がいない" }
  ];
  
  let yPos = 1.9;
  items.forEach((item, i) => {
    addBadge(slide, (i + 1).toString(), 0.5, yPos, 0.45);
    slide.addText(item.title, {
      x: 1.05, y: yPos, w: 2.4, h: 0.45,
      fontSize: 13, fontFace: "Noto Sans JP", bold: true,
      color: DARK_GREEN, align: "left", valign: "middle", margin: 0
    });
    slide.addText(item.desc, {
      x: 3.5, y: yPos, w: 6.0, h: 0.45,
      fontSize: 12, fontFace: "Noto Sans JP",
      color: TEXT_MEDIUM, align: "left", valign: "middle", margin: 0
    });
    yPos += 0.6;
  });
}
```

### パターン4: 2カラム比較スライド

```javascript
{
  let slide = pres.addSlide();
  slide.background = { color: WHITE };
  
  slide.addText("比較タイトル", {
    x: 0.5, y: 0.3, w: 9, h: 0.5,
    fontSize: 20, fontFace: "Noto Sans JP", bold: true,
    color: DARK_GREEN, align: "left", valign: "middle", margin: 0
  });
  
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 0.85, w: 5.5, h: 0.05,
    fill: { color: HIGHLIGHT_YELLOW }
  });
  
  // 左カラム（グレー背景）
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 2.1, w: 4.3, h: 0.5,
    fill: { color: LIGHT_GRAY }
  });
  slide.addText("従来のアプローチ", {
    x: 0.5, y: 2.1, w: 4.3, h: 0.5,
    fontSize: 13, fontFace: "Noto Sans JP", bold: true,
    color: TEXT_MEDIUM, align: "center", valign: "middle", margin: 0
  });
  slide.addText("• ポイント1\n• ポイント2\n• ポイント3", {
    x: 0.7, y: 2.7, w: 3.9, h: 1.5,
    fontSize: 12, fontFace: "Noto Sans JP",
    color: TEXT_MEDIUM, align: "left", valign: "top", margin: 0, lineSpacing: 18
  });
  
  // 右カラム（ダークグリーン背景 - 強調）
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 2.1, w: 4.3, h: 0.5,
    fill: { color: DARK_GREEN }
  });
  slide.addText("新しいアプローチ", {
    x: 5.2, y: 2.1, w: 4.3, h: 0.5,
    fontSize: 13, fontFace: "Noto Sans JP", bold: true,
    color: WHITE, align: "center", valign: "middle", margin: 0
  });
  slide.addText("• ポイント1\n• ポイント2\n• ポイント3", {
    x: 5.4, y: 2.7, w: 3.9, h: 1.5,
    fontSize: 12, fontFace: "Noto Sans JP",
    color: TEXT_DARK, align: "left", valign: "top", margin: 0, lineSpacing: 18
  });
}
```

### パターン5: セクション区切りスライド

```javascript
function addSectionSlide(title) {
  let slide = pres.addSlide();
  slide.background = { color: WHITE };
  
  // 中央にダークグリーンの帯
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 1.8, w: 10, h: 2.0,
    fill: { color: DARK_GREEN }
  });
  
  // タイトル
  slide.addText(title, {
    x: 0.5, y: 1.8, w: 9, h: 2.0,
    fontSize: 36, fontFace: "Noto Sans JP", bold: true,
    color: WHITE, align: "center", valign: "middle"
  });
}

// 使用例
addSectionSlide("解決の方向性");
```

### パターン6: ステートメントスライド（Cynthialyの前提）

```javascript
{
  let slide = pres.addSlide();
  slide.background = { color: DARK_GREEN };
  
  // アクセントライン
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.3, w: 2, h: 0.06,
    fill: { color: CREAM_YELLOW }
  });
  
  // メインメッセージ
  slide.addText("業務プロセスを起点にしない限り、\nAIエージェントは定着しない", {
    x: 0.5, y: 1.5, w: 9, h: 1.5,
    fontSize: 32, fontFace: "Noto Sans JP", bold: true,
    color: WHITE, align: "left", valign: "top", margin: 0, lineSpacing: 44
  });
  
  // サブメッセージ
  slide.addText("── Cynthialyの前提", {
    x: 0.5, y: 3.2, w: 9, h: 0.5,
    fontSize: 16, fontFace: "Noto Sans JP",
    color: CREAM_YELLOW, align: "left", valign: "top", margin: 0
  });
  
  // 説明ボックス
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.0, w: 9, h: 1.2,
    fill: { color: "1A3D38" }
  });
  slide.addText("説明文をここに記載", {
    x: 0.7, y: 4.0, w: 8.6, h: 1.2,
    fontSize: 12, fontFace: "Noto Sans JP",
    color: WHITE, align: "left", valign: "middle", margin: 0
  });
}
```

### パターン7: 3カラム支援内容スライド

```javascript
{
  let slide = pres.addSlide();
  slide.background = { color: WHITE };
  
  slide.addText("提供する3つの支援", {
    x: 0.5, y: 0.3, w: 9, h: 0.5,
    fontSize: 22, fontFace: "Noto Sans JP", bold: true,
    color: DARK_GREEN, align: "left", valign: "middle", margin: 0
  });
  
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 0.85, w: 3.5, h: 0.05,
    fill: { color: HIGHLIGHT_YELLOW }
  });
  
  const supports = [
    { phase: "Step 1", title: "業務棚卸し", desc: "現状業務の可視化" },
    { phase: "Step 2", title: "プロセス設計", desc: "役割分担設計" },
    { phase: "Step 3", title: "プロトタイプ", desc: "検証・改善" }
  ];
  
  supports.forEach((s, i) => {
    const xPos = 0.5 + i * 3.1;
    
    // フェーズラベル
    slide.addShape(pres.shapes.RECTANGLE, {
      x: xPos, y: 1.5, w: 2.9, h: 0.5,
      fill: { color: DARK_GREEN }
    });
    slide.addText(s.phase, {
      x: xPos, y: 1.5, w: 2.9, h: 0.5,
      fontSize: 12, fontFace: "Noto Sans JP", bold: true,
      color: WHITE, align: "center", valign: "middle", margin: 0
    });
    
    // コンテンツボックス
    slide.addShape(pres.shapes.RECTANGLE, {
      x: xPos, y: 2.0, w: 2.9, h: 2.0,
      fill: { color: LIGHT_GRAY }
    });
    slide.addText(s.title, {
      x: xPos, y: 2.1, w: 2.9, h: 0.6,
      fontSize: 13, fontFace: "Noto Sans JP", bold: true,
      color: TEXT_DARK, align: "center", valign: "middle", margin: 0
    });
    slide.addText(s.desc, {
      x: xPos + 0.1, y: 2.8, w: 2.7, h: 1.0,
      fontSize: 11, fontFace: "Noto Sans JP",
      color: TEXT_MEDIUM, align: "center", valign: "top", margin: 0
    });
  });
}
```

### パターン8: まとめスライド（ダークグリーン背景）

```javascript
{
  let slide = pres.addSlide();
  slide.background = { color: DARK_GREEN };
  
  // アクセントライン
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 0.8, w: 2, h: 0.06,
    fill: { color: CREAM_YELLOW }
  });
  
  // タイトル
  slide.addText("まとめ：業務プロセスから始める\nAIエージェント導入を、Cynthialyと一緒に", {
    x: 0.5, y: 1.0, w: 9, h: 1.0,
    fontSize: 24, fontFace: "Noto Sans JP", bold: true,
    color: WHITE, align: "left", valign: "top", margin: 0, lineSpacing: 32
  });
  
  // 箇条書きポイント（クリームイエローバッジ）
  const points = [
    "ポイント1の説明文",
    "ポイント2の説明文",
    "ポイント3の説明文"
  ];
  
  points.forEach((pt, i) => {
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.5, y: 2.2 + i * 0.85, w: 0.4, h: 0.4,
      fill: { color: CREAM_YELLOW }
    });
    slide.addText((i + 1).toString(), {
      x: 0.5, y: 2.2 + i * 0.85, w: 0.4, h: 0.4,
      fontSize: 14, fontFace: "Noto Sans JP", bold: true,
      color: TEXT_DARK, align: "center", valign: "middle", margin: 0
    });
    slide.addText(pt, {
      x: 1.0, y: 2.2 + i * 0.85, w: 8.5, h: 0.75,
      fontSize: 12, fontFace: "Noto Sans JP",
      color: WHITE, align: "left", valign: "middle", margin: 0
    });
  });
  
  // フッターメッセージ
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.8, w: 9, h: 0.6,
    fill: { color: "1A3D38" }
  });
  slide.addText("Cynthialyは、「自社でエージェントを増やし・育て続けられる状態」をゴールに伴走支援を提供する。", {
    x: 0.7, y: 4.8, w: 8.6, h: 0.6,
    fontSize: 11, fontFace: "Noto Sans JP",
    color: CREAM_YELLOW, align: "center", valign: "middle", margin: 0
  });
}
```

## ボトムインサイトボックス

多くのスライドで使用する「ダークグリーン背景のメッセージボックス」

```javascript
// スライド下部に配置
slide.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 3.5, w: 9, h: 1.5,
  fill: { color: DARK_GREEN }
});
slide.addText("重要なメッセージをここに記載", {
  x: 0.7, y: 3.5, w: 8.6, h: 1.5,
  fontSize: 16, fontFace: "Noto Sans JP", bold: true,
  color: WHITE, align: "center", valign: "middle", margin: 0
});
```

## 基本テンプレート

```javascript
const pptxgen = require("pptxgenjs");

// カラーパレット
const DARK_GREEN = "0D2623";
const CREAM_YELLOW = "E8DE9F";
const LIGHT_GRAY = "F3F3F3";
const TEXT_DARK = "0D2623";
const TEXT_MEDIUM = "434343";
const TEXT_LIGHT = "595959";
const WHITE = "FFFFFF";
const HIGHLIGHT_YELLOW = "FFFF00";

// プレゼンテーション作成
let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Cynthialy株式会社";
pres.title = "プレゼンテーションタイトル";

// ヘルパー関数
function addSectionSlide(title) {
  let slide = pres.addSlide();
  slide.background = { color: WHITE };
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 1.8, w: 10, h: 2.0,
    fill: { color: DARK_GREEN }
  });
  slide.addText(title, {
    x: 0.5, y: 1.8, w: 9, h: 2.0,
    fontSize: 36, fontFace: "Noto Sans JP", bold: true,
    color: WHITE, align: "center", valign: "middle"
  });
}

function addBadge(slide, num, x, y, size = 0.5) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: size, h: size,
    fill: { color: DARK_GREEN }
  });
  slide.addText(String(num), {
    x: x, y: y, w: size, h: size,
    fontSize: 18, fontFace: "Noto Sans JP", bold: true,
    color: WHITE, align: "center", valign: "middle", margin: 0
  });
}

function addStatBox(slide, stat, label, x, y, w, h) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: x, y: y, w: w, h: h,
    fill: { color: LIGHT_GRAY }
  });
  slide.addText(stat, {
    x: x, y: y + 0.15, w: w, h: h * 0.5,
    fontSize: 36, fontFace: "Noto Sans JP", bold: true,
    color: DARK_GREEN, align: "center", valign: "middle", margin: 0
  });
  slide.addText(label, {
    x: x + 0.1, y: y + h * 0.55, w: w - 0.2, h: h * 0.4,
    fontSize: 11, fontFace: "Noto Sans JP",
    color: TEXT_MEDIUM, align: "center", valign: "top", margin: 0
  });
}

// スライドを追加...

// 保存
pres.writeFile({ fileName: "/home/claude/output.pptx" })
  .then(() => console.log("Presentation saved"))
  .catch(err => console.error(err));
```

## デザインガイドライン

### マージンとスペーシング
- 左右マージン: 0.5インチ
- 上マージン: 0.3インチ
- 下マージン: 0.2インチ
- タイトルと本文の間隔: 0.5-0.7インチ
- リスト項目間隔: 0.6インチ

### フォントサイズ
- タイトルスライド: 36-40pt (bold)
- スライドタイトル: 18-22pt (bold)
- キーメッセージ: 14-16pt (bold)
- 本文: 11-13pt
- 出典・補足: 9pt

### 色使いのルール
1. **ダークグリーン背景**: タイトルスライド、ステートメントスライド、まとめスライド、ボトムインサイトボックス
2. **白背景**: 通常のコンテンツスライド
3. **クリームイエロー**: アクセントライン、キーメッセージボックス、強調バッジ
4. **黄色**: タイトル下のアンダーライン
5. **ライトグレー**: コンテンツボックス、統計ボックス

## ベストプラクティス

1. **一貫性**: 同じパターンのスライドは同じレイアウトを使用
2. **コントラスト**: ダークグリーン背景には白テキスト、白背景にはダークグリーンテキスト
3. **余白**: スライドを詰め込みすぎない
4. **階層**: タイトル → キーメッセージ → 本文 → 出典の順で視覚的階層を作る
5. **強調**: 最も伝えたいメッセージはダークグリーン背景のボックスに配置


## チャート用 拡張カラーパレット

データビジュアライゼーション用の追加カラー。既存パレットと併用する。

```javascript
const CHART_BLUE = "2D5EA3";       // メインチャート色
const CHART_BLUE_LIGHT = "6699CC"; // セカンダリチャート色
const CHART_BLUE_PALE = "A8C4E0";  // 第3チャート色
const CHART_RED = "C0392B";        // 減少・警告
const CHART_ORANGE = "E67E22";     // アクセント
const CHART_TEAL = "1ABC9C";       // 補助色
const MED_GRAY = "888888";         // 目盛ラベル
const BORDER_GRAY = "CCCCCC";      // グリッドライン
const VERY_LIGHT_GRAY = "F5F5F5";  // チャート背景
```

## チャートパターン

pptxgenjsの図形描画のみで構成。外部チャートライブラリ不要。
すべて `drawXxxChart(slide, data, options)` の統一インターフェースで呼び出す。

### 共通ヘルパー: スライドヘッダー

チャートスライドの共通ヘッダー（タイトル + 黄色アンダーライン）

```javascript
function addSlideHeader(slide, title) {
  slide.addText(title, {
    x: 0.5, y: 0.3, w: 9, h: 0.5,
    fontSize: 22, fontFace: "Noto Sans JP", bold: true,
    color: DARK_GREEN, align: "left", valign: "middle", margin: 0
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 0.85, w: 5.5, h: 0.05,
    fill: { color: HIGHLIGHT_YELLOW }
  });
}
```

### パターン9: 縦棒グラフ (Vertical Bar Chart)

時系列データ、カテゴリ比較、成長トレンドの可視化に使用

```javascript
function drawBarChart(slide, data, options = {}) {
  const {
    chartX = 0.5, chartY = 1.2,
    chartW = 9.0, chartH = 3.2,
    unit = "", showValues = true, barGap = 0.08,
    defaultColor = DARK_GREEN, axisColor = TEXT_MEDIUM,
    gridLines = 4, title = null, source = null
  } = options;

  if (title) {
    slide.addText(title, {
      x: chartX, y: chartY - 0.35, w: chartW, h: 0.3,
      fontSize: 11, fontFace: "Noto Sans JP", bold: true,
      color: TEXT_DARK, align: "left", valign: "middle", margin: 0
    });
  }

  const n = data.length;
  const maxVal = Math.max(...data.map(d => d.value));
  const barW = (chartW - (n + 1) * barGap) / n;
  const bottom = chartY + chartH;

  for (let i = 0; i <= gridLines; i++) {
    const gY = chartY + (chartH / gridLines) * i;
    const gVal = maxVal * (1 - i / gridLines);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: chartX, y: gY, w: chartW, h: 0.005,
      fill: { color: BORDER_GRAY }
    });
    slide.addText(Math.round(gVal).toLocaleString() + unit, {
      x: chartX - 0.8, y: gY - 0.1, w: 0.75, h: 0.2,
      fontSize: 7, fontFace: "Noto Sans JP",
      color: MED_GRAY, align: "right", valign: "middle", margin: 0
    });
  }

  slide.addShape(pres.shapes.RECTANGLE, {
    x: chartX, y: bottom, w: chartW, h: 0.015,
    fill: { color: axisColor }
  });

  data.forEach((d, i) => {
    const barH = (d.value / maxVal) * chartH;
    const x = chartX + barGap + i * (barW + barGap);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: x, y: bottom - barH, w: barW, h: barH,
      fill: { color: d.color || defaultColor }
    });
    if (showValues) {
      slide.addText(d.value.toLocaleString() + unit, {
        x: x - 0.1, y: bottom - barH - 0.22, w: barW + 0.2, h: 0.2,
        fontSize: 8, fontFace: "Noto Sans JP", bold: true,
        color: TEXT_DARK, align: "center", valign: "middle", margin: 0
      });
    }
    slide.addText(d.label, {
      x: x - 0.05, y: bottom + 0.04, w: barW + 0.1, h: 0.25,
      fontSize: 7.5, fontFace: "Noto Sans JP",
      color: TEXT_MEDIUM, align: "center", valign: "top", margin: 0
    });
  });

  if (source) {
    slide.addText(source, {
      x: chartX, y: bottom + 0.35, w: chartW, h: 0.2,
      fontSize: 7, fontFace: "Noto Sans JP",
      color: TEXT_LIGHT, align: "left", valign: "middle", margin: 0
    });
  }
}

// 使用例
drawBarChart(slide, [
  { label: "2023", value: 420, color: CHART_BLUE_LIGHT },
  { label: "2024", value: 615, color: CHART_BLUE },
  { label: "2025E", value: 860, color: DARK_GREEN }
], {
  chartX: 1.3, chartY: 1.3, chartW: 8.0, chartH: 3.0,
  title: "市場規模推移 ($M)",
  source: "出典：〇〇調査",
  gridLines: 5
});
```

### パターン10: 積み上げ棒グラフ (Stacked Bar Chart)

構成比の変化、セグメント内訳の時系列比較に使用

```javascript
function drawStackedBarChart(slide, categories, series, options = {}) {
  const {
    chartX = 0.5, chartY = 1.2,
    chartW = 6.0, chartH = 3.2,
    unit = "", barGap = 0.12, axisColor = TEXT_MEDIUM,
    gridLines = 4, title = null, source = null,
    showLegend = true, legendX = null, legendY = null
  } = options;

  if (title) {
    slide.addText(title, {
      x: chartX, y: chartY - 0.35, w: chartW, h: 0.3,
      fontSize: 11, fontFace: "Noto Sans JP", bold: true,
      color: TEXT_DARK, align: "left", valign: "middle", margin: 0
    });
  }

  const n = categories.length;
  const barW = (chartW - (n + 1) * barGap) / n;
  const bottom = chartY + chartH;
  const totals = categories.map((_, ci) =>
    series.reduce((sum, s) => sum + s.values[ci], 0)
  );
  const maxVal = Math.max(...totals);

  for (let i = 0; i <= gridLines; i++) {
    const gY = chartY + (chartH / gridLines) * i;
    const gVal = maxVal * (1 - i / gridLines);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: chartX, y: gY, w: chartW, h: 0.005,
      fill: { color: BORDER_GRAY }
    });
    slide.addText(Math.round(gVal).toLocaleString() + unit, {
      x: chartX - 0.8, y: gY - 0.1, w: 0.75, h: 0.2,
      fontSize: 7, fontFace: "Noto Sans JP",
      color: MED_GRAY, align: "right", valign: "middle", margin: 0
    });
  }

  slide.addShape(pres.shapes.RECTANGLE, {
    x: chartX, y: bottom, w: chartW, h: 0.015,
    fill: { color: axisColor }
  });

  categories.forEach((cat, ci) => {
    const x = chartX + barGap + ci * (barW + barGap);
    let currentY = bottom;
    series.forEach((s) => {
      const segH = (s.values[ci] / maxVal) * chartH;
      currentY -= segH;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: x, y: currentY, w: barW, h: segH,
        fill: { color: s.color }
      });
    });
    const totalH = (totals[ci] / maxVal) * chartH;
    slide.addText(totals[ci].toLocaleString() + unit, {
      x: x - 0.1, y: bottom - totalH - 0.22, w: barW + 0.2, h: 0.2,
      fontSize: 8, fontFace: "Noto Sans JP", bold: true,
      color: TEXT_DARK, align: "center", valign: "middle", margin: 0
    });
    slide.addText(cat, {
      x: x - 0.05, y: bottom + 0.04, w: barW + 0.1, h: 0.25,
      fontSize: 7.5, fontFace: "Noto Sans JP",
      color: TEXT_MEDIUM, align: "center", valign: "top", margin: 0
    });
  });

  if (showLegend) {
    const lx = legendX || (chartX + chartW + 0.3);
    const ly = legendY || chartY;
    series.forEach((s, i) => {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: lx, y: ly + i * 0.32, w: 0.2, h: 0.16,
        fill: { color: s.color }
      });
      slide.addText(s.name, {
        x: lx + 0.28, y: ly + i * 0.32 - 0.02, w: 2.5, h: 0.2,
        fontSize: 8, fontFace: "Noto Sans JP",
        color: TEXT_MEDIUM, align: "left", valign: "middle", margin: 0
      });
    });
  }

  if (source) {
    slide.addText(source, {
      x: chartX, y: bottom + 0.35, w: chartW + 3, h: 0.2,
      fontSize: 7, fontFace: "Noto Sans JP",
      color: TEXT_LIGHT, align: "left", valign: "middle", margin: 0
    });
  }
}

// 使用例
drawStackedBarChart(slide,
  ["2023", "2024", "2025E"],
  [
    { name: "セグメントA", values: [200, 310, 480], color: DARK_GREEN },
    { name: "セグメントB", values: [100, 160, 250], color: CHART_BLUE },
    { name: "セグメントC", values: [40, 80, 130], color: CREAM_YELLOW }
  ], {
    chartX: 1.3, chartY: 1.3, chartW: 6.0, chartH: 3.0,
    title: "セグメント別市場構成 ($M)"
  }
);
```

### パターン11: 横棒グラフ (Horizontal Bar Chart)

ランキング、スコア比較、ベンチマーク結果の表示に使用

```javascript
function drawHorizontalBarChart(slide, data, options = {}) {
  const {
    chartX = 2.0, chartY = 1.2,
    chartW = 6.5, chartH = 3.6,
    maxValue = null, unit = "",
    defaultColor = DARK_GREEN, bgColor = VERY_LIGHT_GRAY,
    title = null, source = null,
    showValues = true, barH = null, barGap = 0.06
  } = options;

  if (title) {
    slide.addText(title, {
      x: 0.5, y: chartY - 0.35, w: chartW + 2, h: 0.3,
      fontSize: 11, fontFace: "Noto Sans JP", bold: true,
      color: TEXT_DARK, align: "left", valign: "middle", margin: 0
    });
  }

  const n = data.length;
  const effectiveBarH = barH || Math.min(0.4, (chartH - (n - 1) * barGap) / n);
  const totalHeight = n * effectiveBarH + (n - 1) * barGap;
  const startY = chartY + (chartH - totalHeight) / 2;
  const mxVal = maxValue || Math.max(...data.map(d => d.value));

  data.forEach((d, i) => {
    const y = startY + i * (effectiveBarH + barGap);
    const barWidth = (d.value / mxVal) * chartW;

    slide.addText(d.label, {
      x: chartX - 1.6, y: y, w: 1.5, h: effectiveBarH,
      fontSize: 9, fontFace: "Noto Sans JP", bold: d.highlight ? true : false,
      color: d.highlight ? DARK_GREEN : TEXT_MEDIUM,
      align: "right", valign: "middle", margin: 0
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: chartX, y: y, w: chartW, h: effectiveBarH,
      fill: { color: bgColor }
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: chartX, y: y, w: barWidth, h: effectiveBarH,
      fill: { color: d.color || defaultColor }
    });
    if (showValues) {
      slide.addText(d.value.toFixed(1) + unit, {
        x: chartX + barWidth + 0.08, y: y, w: 0.7, h: effectiveBarH,
        fontSize: 8, fontFace: "Noto Sans JP", bold: true,
        color: d.color || defaultColor,
        align: "left", valign: "middle", margin: 0
      });
    }
  });

  if (source) {
    slide.addText(source, {
      x: 0.5, y: chartY + chartH + 0.15, w: 9, h: 0.2,
      fontSize: 7, fontFace: "Noto Sans JP",
      color: TEXT_LIGHT, align: "left", valign: "middle", margin: 0
    });
  }
}

// 使用例
drawHorizontalBarChart(slide, [
  { label: "モデルA", value: 9.2, color: CHART_RED, highlight: true },
  { label: "モデルB", value: 8.5, color: CHART_BLUE },
  { label: "モデルC", value: 7.8, color: MED_GRAY }
], {
  chartX: 2.3, chartY: 1.3, chartW: 5.5, chartH: 3.5,
  maxValue: 10,
  title: "ベンチマークスコア比較 (10点満点)",
  barH: 0.35, barGap: 0.08
});
```

### パターン12: ウォーターフォールチャート (Waterfall Chart)

増減分析、コスト構造、ブリッジ分析に使用。
各バーの type で start(開始値)/add(増加)/subtract(減少)/total(合計) を指定。

```javascript
function drawWaterfallChart(slide, data, options = {}) {
  const {
    chartX = 0.5, chartY = 1.2,
    chartW = 9.0, chartH = 3.2,
    unit = "", addColor = DARK_GREEN,
    subtractColor = CHART_RED, totalColor = CHART_BLUE,
    title = null, source = null,
    gridLines = 4, barGap = 0.1
  } = options;

  if (title) {
    slide.addText(title, {
      x: chartX, y: chartY - 0.35, w: chartW, h: 0.3,
      fontSize: 11, fontFace: "Noto Sans JP", bold: true,
      color: TEXT_DARK, align: "left", valign: "middle", margin: 0
    });
  }

  const n = data.length;
  const barW = (chartW - (n + 1) * barGap) / n;
  const bottom = chartY + chartH;

  let runningTotal = 0;
  let allValues = [];
  data.forEach(d => {
    if (d.type === "start" || d.type === "total") runningTotal = d.value;
    else if (d.type === "add") runningTotal += d.value;
    else if (d.type === "subtract") runningTotal -= d.value;
    allValues.push(runningTotal);
  });
  const maxVal = Math.max(...allValues, ...data.map(d => d.value)) * 1.1;

  for (let i = 0; i <= gridLines; i++) {
    const gY = chartY + (chartH / gridLines) * i;
    const gVal = maxVal * (1 - i / gridLines);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: chartX, y: gY, w: chartW, h: 0.005,
      fill: { color: BORDER_GRAY }
    });
    slide.addText(Math.round(gVal).toLocaleString() + unit, {
      x: chartX - 0.8, y: gY - 0.1, w: 0.75, h: 0.2,
      fontSize: 7, fontFace: "Noto Sans JP",
      color: MED_GRAY, align: "right", valign: "middle", margin: 0
    });
  }

  slide.addShape(pres.shapes.RECTANGLE, {
    x: chartX, y: bottom, w: chartW, h: 0.015,
    fill: { color: TEXT_MEDIUM }
  });

  let cumulative = 0;
  data.forEach((d, i) => {
    const x = chartX + barGap + i * (barW + barGap);
    let barTop, barHeight, color;

    if (d.type === "start" || d.type === "total") {
      barHeight = (d.value / maxVal) * chartH;
      barTop = bottom - barHeight;
      color = totalColor;
      cumulative = d.value;
    } else if (d.type === "add") {
      const prevTop = bottom - (cumulative / maxVal) * chartH;
      barHeight = (d.value / maxVal) * chartH;
      barTop = prevTop - barHeight;
      color = addColor;
      cumulative += d.value;
    } else {
      barHeight = (d.value / maxVal) * chartH;
      barTop = bottom - (cumulative / maxVal) * chartH;
      color = subtractColor;
      cumulative -= d.value;
    }

    slide.addShape(pres.shapes.RECTANGLE, {
      x: x, y: barTop, w: barW, h: barHeight,
      fill: { color: color }
    });

    const sign = d.type === "subtract" ? "-" : (d.type === "add" ? "+" : "");
    slide.addText(sign + d.value.toLocaleString() + unit, {
      x: x - 0.1, y: barTop - 0.22, w: barW + 0.2, h: 0.2,
      fontSize: 7.5, fontFace: "Noto Sans JP", bold: true,
      color: color, align: "center", valign: "middle", margin: 0
    });
    slide.addText(d.label, {
      x: x - 0.1, y: bottom + 0.04, w: barW + 0.2, h: 0.3,
      fontSize: 7, fontFace: "Noto Sans JP",
      color: TEXT_MEDIUM, align: "center", valign: "top", margin: 0
    });

    // コネクターライン（次のバーへ、totalの前以外）
    if (i < n - 1 && data[i + 1].type !== "total") {
      const connectorY = bottom - (cumulative / maxVal) * chartH;
      const nextX = chartX + barGap + (i + 1) * (barW + barGap);
      slide.addShape(pres.shapes.RECTANGLE, {
        x: x + barW, y: connectorY, w: nextX - (x + barW), h: 0.008,
        fill: { color: MED_GRAY }
      });
    }
  });

  if (source) {
    slide.addText(source, {
      x: chartX, y: bottom + 0.4, w: chartW, h: 0.2,
      fontSize: 7, fontFace: "Noto Sans JP",
      color: TEXT_LIGHT, align: "left", valign: "middle", margin: 0
    });
  }
}

// 使用例
drawWaterfallChart(slide, [
  { label: "開始値", value: 615, type: "start" },
  { label: "成長A", value: 120, type: "add" },
  { label: "成長B", value: 65, type: "add" },
  { label: "減少C", value: 83, type: "subtract" },
  { label: "最終値", value: 717, type: "total" }
], {
  chartX: 1.3, chartY: 1.3, chartW: 8.0, chartH: 3.0,
  title: "ブリッジ分析 ($M)"
});
```

### パターン13: バブルチャート (Bubble Plot / 2×2マトリクス)

競争ポジショニング、2軸評価マップに使用。
quadrantLabels を指定すると4象限ラベル付き2×2マトリクスになる。

```javascript
function drawBubblePlot(slide, data, options = {}) {
  const {
    plotX = 1.3, plotY = 1.2,
    plotW = 5.5, plotH = 3.5,
    xLabel = "X軸", yLabel = "Y軸",
    xRange = [0, 10], yRange = [0, 10],
    defaultColor = DARK_GREEN,
    title = null, source = null,
    quadrantLabels = null  // ["左下", "右下", "左上", "右上"]
  } = options;

  if (title) {
    slide.addText(title, {
      x: 0.5, y: plotY - 0.4, w: 9, h: 0.3,
      fontSize: 11, fontFace: "Noto Sans JP", bold: true,
      color: TEXT_DARK, align: "left", valign: "middle", margin: 0
    });
  }

  // 四分割背景（オプション）
  if (quadrantLabels) {
    const qColors = [VERY_LIGHT_GRAY, "EDF2F7", "EDF2F7", "E2EBD6"];
    const qPos = [
      [plotX, plotY + plotH / 2, plotW / 2, plotH / 2],
      [plotX + plotW / 2, plotY + plotH / 2, plotW / 2, plotH / 2],
      [plotX, plotY, plotW / 2, plotH / 2],
      [plotX + plotW / 2, plotY, plotW / 2, plotH / 2]
    ];
    qPos.forEach((q, i) => {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: q[0], y: q[1], w: q[2], h: q[3],
        fill: { color: qColors[i] }
      });
      slide.addText(quadrantLabels[i], {
        x: q[0] + 0.08, y: q[1] + 0.05, w: q[2] - 0.16, h: 0.2,
        fontSize: 7, fontFace: "Noto Sans JP", italic: true,
        color: MED_GRAY, align: "center", valign: "top", margin: 0
      });
    });
  }

  // 軸線
  slide.addShape(pres.shapes.RECTANGLE, {
    x: plotX, y: plotY + plotH / 2, w: plotW, h: 0.012,
    fill: { color: TEXT_MEDIUM }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: plotX + plotW / 2, y: plotY, w: 0.012, h: plotH,
    fill: { color: TEXT_MEDIUM }
  });

  // 軸ラベル
  slide.addText(xLabel + " →", {
    x: plotX, y: plotY + plotH + 0.08, w: plotW, h: 0.22,
    fontSize: 8, fontFace: "Noto Sans JP", bold: true,
    color: TEXT_MEDIUM, align: "center", valign: "top", margin: 0
  });

  // バブル描画
  data.forEach(d => {
    const cx = plotX + ((d.x - xRange[0]) / (xRange[1] - xRange[0])) * plotW;
    const cy = plotY + plotH - ((d.y - yRange[0]) / (yRange[1] - yRange[0])) * plotH;
    const sz = d.size || 0.35;
    slide.addShape(pres.shapes.OVAL, {
      x: cx - sz / 2, y: cy - sz / 2, w: sz, h: sz,
      fill: { color: d.color || defaultColor, transparency: 25 },
      line: { color: d.color || defaultColor, width: 1 }
    });
    slide.addText(d.label, {
      x: cx - 0.6, y: cy + sz / 2 + 0.02, w: 1.2, h: 0.18,
      fontSize: 7, fontFace: "Noto Sans JP", bold: true,
      color: d.color || defaultColor,
      align: "center", valign: "top", margin: 0
    });
  });

  if (source) {
    slide.addText(source, {
      x: 0.5, y: plotY + plotH + 0.35, w: 9, h: 0.2,
      fontSize: 7, fontFace: "Noto Sans JP",
      color: TEXT_LIGHT, align: "left", valign: "middle", margin: 0
    });
  }
}

// 使用例
drawBubblePlot(slide, [
  { label: "企業A", x: 8.5, y: 7.5, size: 0.42, color: CHART_RED },
  { label: "企業B", x: 7.0, y: 9.0, size: 0.40, color: DARK_GREEN }
], {
  plotX: 1.5, plotY: 1.3, plotW: 5.5, plotH: 3.5,
  xLabel: "技術力", yLabel: "市場規模",
  xRange: [3, 10], yRange: [3, 10],
  title: "ポジショニングマップ",
  quadrantLabels: ["Emerging", "Challenger", "Niche", "Leader"]
});
```

### パターン14: 折れ線グラフ (Line Chart)

トレンド推移の比較、時系列の複数指標追跡に使用。
pptxgenjsの小さな円形図形の点列でラインを近似描画する。

```javascript
function drawLineChart(slide, categories, series, options = {}) {
  const {
    chartX = 1.3, chartY = 1.2,
    chartW = 7.0, chartH = 3.2,
    unit = "", title = null, source = null,
    gridLines = 4, showDots = true,
    showLegend = true, legendX = null, legendY = null
  } = options;

  if (title) {
    slide.addText(title, {
      x: 0.5, y: chartY - 0.35, w: 9, h: 0.3,
      fontSize: 11, fontFace: "Noto Sans JP", bold: true,
      color: TEXT_DARK, align: "left", valign: "middle", margin: 0
    });
  }

  const bottom = chartY + chartH;
  const allVals = series.flatMap(s => s.values);
  const maxVal = Math.max(...allVals) * 1.1;
  const n = categories.length;
  const stepX = chartW / (n - 1);

  // グリッドライン
  for (let i = 0; i <= gridLines; i++) {
    const gY = chartY + (chartH / gridLines) * i;
    const gVal = maxVal * (1 - i / gridLines);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: chartX, y: gY, w: chartW, h: 0.005,
      fill: { color: BORDER_GRAY }
    });
    slide.addText(Math.round(gVal).toLocaleString() + unit, {
      x: chartX - 0.8, y: gY - 0.1, w: 0.75, h: 0.2,
      fontSize: 7, fontFace: "Noto Sans JP",
      color: MED_GRAY, align: "right", valign: "middle", margin: 0
    });
  }

  // ベースライン
  slide.addShape(pres.shapes.RECTANGLE, {
    x: chartX, y: bottom, w: chartW, h: 0.015,
    fill: { color: TEXT_MEDIUM }
  });

  // X軸ラベル
  categories.forEach((cat, i) => {
    slide.addText(cat, {
      x: chartX + i * stepX - 0.3, y: bottom + 0.04, w: 0.6, h: 0.25,
      fontSize: 7, fontFace: "Noto Sans JP",
      color: TEXT_MEDIUM, align: "center", valign: "top", margin: 0
    });
  });

  // ラインを点列で近似描画
  series.forEach(s => {
    for (let i = 0; i < n - 1; i++) {
      const x1 = chartX + i * stepX;
      const y1 = bottom - (s.values[i] / maxVal) * chartH;
      const x2 = chartX + (i + 1) * stepX;
      const y2 = bottom - (s.values[i + 1] / maxVal) * chartH;
      const dx = x2 - x1, dy = y2 - y1;
      const len = Math.sqrt(dx * dx + dy * dy);
      const steps = Math.max(Math.round(len / 0.03), 5);
      for (let st = 0; st < steps; st++) {
        const t = st / steps;
        slide.addShape(pres.shapes.OVAL, {
          x: x1 + t * dx - 0.012, y: y1 + t * dy - 0.012,
          w: 0.024, h: 0.024,
          fill: { color: s.color }
        });
      }
    }

    // データポイント（ドット）
    if (showDots) {
      s.values.forEach((v, i) => {
        slide.addShape(pres.shapes.OVAL, {
          x: chartX + i * stepX - 0.06,
          y: bottom - (v / maxVal) * chartH - 0.06,
          w: 0.12, h: 0.12,
          fill: { color: s.color },
          line: { color: WHITE, width: 1.5 }
        });
      });
    }

    // 最終値ラベル
    const lastVal = s.values[n - 1];
    slide.addText(lastVal.toLocaleString() + unit, {
      x: chartX + (n - 1) * stepX + 0.1,
      y: bottom - (lastVal / maxVal) * chartH - 0.1,
      w: 0.7, h: 0.2,
      fontSize: 7, fontFace: "Noto Sans JP", bold: true,
      color: s.color, align: "left", valign: "middle", margin: 0
    });
  });

  // 凡例
  if (showLegend) {
    const lx = legendX || (chartX + chartW + 0.3);
    const ly = legendY || chartY;
    series.forEach((s, i) => {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: lx, y: ly + i * 0.3, w: 0.25, h: 0.04,
        fill: { color: s.color }
      });
      slide.addText(s.name, {
        x: lx + 0.32, y: ly + i * 0.3 - 0.08, w: 1.8, h: 0.2,
        fontSize: 8, fontFace: "Noto Sans JP",
        color: TEXT_MEDIUM, align: "left", valign: "middle", margin: 0
      });
    });
  }

  if (source) {
    slide.addText(source, {
      x: 0.5, y: bottom + 0.35, w: 9, h: 0.2,
      fontSize: 7, fontFace: "Noto Sans JP",
      color: TEXT_LIGHT, align: "left", valign: "middle", margin: 0
    });
  }
}

// 使用例
drawLineChart(slide,
  ["Q1", "Q2", "Q3", "Q4"],
  [
    { name: "製品A", values: [15, 38, 68, 110], color: CHART_BLUE },
    { name: "製品B", values: [10, 25, 45, 80], color: CHART_TEAL }
  ], {
    chartX: 1.3, chartY: 1.3, chartW: 6.5, chartH: 3.0,
    title: "四半期推移比較"
  }
);
```

### パターン15: プログレスバー付きスコアカード

複数項目 × 複数対象のスコア比較、成熟度評価に使用

```javascript
function drawProgressComparison(slide, data, options = {}) {
  const {
    startX = 0.5, startY = 1.2,
    width = 9.0, rowH = 0.45, maxValue = 10,
    title = null, source = null,
    barHeight = 0.16, bgColor = VERY_LIGHT_GRAY,
    categories = null  // ヘッダー行のモデル/企業名
  } = options;

  if (title) {
    slide.addText(title, {
      x: startX, y: startY - 0.35, w: width, h: 0.3,
      fontSize: 11, fontFace: "Noto Sans JP", bold: true,
      color: TEXT_DARK, align: "left", valign: "middle", margin: 0
    });
  }

  const labelW = 1.6;
  const barAreaW = width - labelW - 0.1;

  // ヘッダー行
  if (categories) {
    slide.addText("評価項目", {
      x: startX, y: startY, w: labelW, h: 0.3,
      fontSize: 8, fontFace: "Noto Sans JP", bold: true,
      color: TEXT_MEDIUM, align: "left", valign: "middle", margin: 0
    });
    const catW = barAreaW / categories.length;
    categories.forEach((cat, i) => {
      slide.addText(cat, {
        x: startX + labelW + i * catW, y: startY, w: catW, h: 0.3,
        fontSize: 8, fontFace: "Noto Sans JP", bold: true,
        color: TEXT_MEDIUM, align: "center", valign: "middle", margin: 0
      });
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: startX, y: startY + 0.3, w: width, h: 0.01,
      fill: { color: BORDER_GRAY }
    });
  }

  const headerOffset = categories ? 0.38 : 0;

  // データ行
  data.forEach((row, ri) => {
    const y = startY + headerOffset + ri * rowH;

    // 偶数行背景
    if (ri % 2 === 0) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: startX, y: y, w: width, h: rowH,
        fill: { color: VERY_LIGHT_GRAY }
      });
    }

    // 行ラベル
    slide.addText(row.label, {
      x: startX + 0.05, y: y, w: labelW - 0.1, h: rowH,
      fontSize: 8.5, fontFace: "Noto Sans JP", bold: true,
      color: TEXT_DARK, align: "left", valign: "middle", margin: 0
    });

    // 各値のプログレスバー
    const catW = barAreaW / row.values.length;
    row.values.forEach((v, vi) => {
      const bx = startX + labelW + vi * catW + catW * 0.1;
      const bw = catW * 0.6;
      const by = y + (rowH - barHeight) / 2;

      // 背景トラック
      slide.addShape(pres.shapes.RECTANGLE, {
        x: bx, y: by, w: bw, h: barHeight,
        fill: { color: "E8E8E8" }
      });
      // 値バー
      slide.addShape(pres.shapes.RECTANGLE, {
        x: bx, y: by, w: (v.score / maxValue) * bw, h: barHeight,
        fill: { color: v.color || DARK_GREEN }
      });
      // スコアラベル
      slide.addText(v.score.toFixed(1), {
        x: bx + bw + 0.03, y: y, w: catW * 0.2, h: rowH,
        fontSize: 7, fontFace: "Noto Sans JP", bold: true,
        color: v.color || TEXT_DARK,
        align: "left", valign: "middle", margin: 0
      });
    });
  });

  if (source) {
    const totalY = startY + headerOffset + data.length * rowH + 0.15;
    slide.addText(source, {
      x: startX, y: totalY, w: width, h: 0.2,
      fontSize: 7, fontFace: "Noto Sans JP",
      color: TEXT_LIGHT, align: "left", valign: "middle", margin: 0
    });
  }
}

// 使用例
drawProgressComparison(slide, [
  { label: "品質", values: [
    { score: 9.2, color: CHART_RED },
    { score: 8.5, color: CHART_BLUE }
  ]},
  { label: "速度", values: [
    { score: 7.8, color: CHART_RED },
    { score: 9.0, color: CHART_BLUE }
  ]}
], {
  startX: 0.3, startY: 1.2, width: 9.4,
  title: "スコアカード比較",
  categories: ["モデルA", "モデルB"]
});
```

## チャートパターン デザインガイドライン

### 共通描画ルール

| 要素 | 仕様 |
|---|---|
| Y軸ラベル | chartX - 0.8、fontSize 7、右寄せ |
| グリッドライン | BORDER_GRAY ("CCCCCC")、高さ 0.005 |
| ベースライン（X軸） | TEXT_MEDIUM、高さ 0.015 |
| 値ラベル | fontSize 8、bold、バーの上または右に配置 |
| カテゴリラベル | fontSize 7.5、TEXT_MEDIUM、ベースライン下 0.04 |
| チャートタイトル | chartY - 0.35、fontSize 11、bold |
| 出典 | fontSize 7、TEXT_LIGHT、グラフ下 0.35〜0.4 |

### チャートカラー使い分け

| 意味 | カラー |
|---|---|
| 増加・ポジティブ | DARK_GREEN ("0D2623") |
| 減少・ネガティブ | CHART_RED ("C0392B") |
| 合計・基準 | CHART_BLUE ("2D5EA3") |
| 実績値 → 予測値 | 濃い色 → 薄い色でグラデーション |
| ハイライト | CREAM_YELLOW ("E8DE9F") |

### 推奨レイアウト配置

| パターン | chartX | chartW | 備考 |
|---|---|---|---|
| グラフ単独 | 1.3 | 8.0 | Y軸ラベル用に左マージン確保 |
| グラフ + 右テーブル | 1.3 | 5.5〜6.0 | 右側にテーブルや凡例を配置 |
| 上下2段構成 | 1.3 | 8.0 | chartH=2.5 を2つ上下に配置 |

### チャートパターン一覧（パターン9〜15）

| # | パターン名 | 関数名 | 主な用途 |
|---|---|---|---|
| 9 | 縦棒グラフ | `drawBarChart` | 時系列、カテゴリ比較 |
| 10 | 積み上げ棒グラフ | `drawStackedBarChart` | 構成比変化、セグメント内訳 |
| 11 | 横棒グラフ | `drawHorizontalBarChart` | ランキング、スコア比較 |
| 12 | ウォーターフォール | `drawWaterfallChart` | 増減分析、ブリッジ |
| 13 | バブルチャート | `drawBubblePlot` | ポジショニング、2軸評価 |
| 14 | 折れ線グラフ | `drawLineChart` | トレンド比較、時系列 |
| 15 | スコアカード | `drawProgressComparison` | 多次元スコア比較 |
