---
name: pptx-compare
description: Compare two PPTX files using a 2-tier XML-based analysis pipeline. slide-compare-lite identifies WHAT changed (with file metadata and summary); slide-compare reveals HOW at the raw XML level.
license: Apache-2.0
---

# PPTX Compare

2層構造のXMLベース比較スキル。全層がXML解析をベースとし、シームレスに連携する。

## 方式の選択

|            | slide-compare-lite                                 | slide-compare                   |
| ---------- | -------------------------------------------------- | ------------------------------- |
| **目的**   | 何が変わったか把握                                 | 詳細をXMLレベルで確認           |
| **正規化** | 強（PowerPoint由来ノイズ全除去）                   | 最小限（属性ソート・空白のみ）  |
| **出力**   | ファイルメタ情報・サマリー・正規化済みunified diff | ノードベースdiff or 直接読み    |
| **速度**   | 速い（CRC32で未変更パートをスキップ）              | 遅い（全パートを4回読んでdiff） |

＊ここでいう「PowerPoint由来ノイズ」とは、ユーザーが編集していないのに保存のたびに書き換わる
XMLのことを指す。言語属性（`lang`/`altLang`）、不可視のテーブル罫線、チャートの書式ヒント、
浮動小数点の表記ゆれ、テキストランの分割・結合などが該当する。
lite はこれらを除去してから比較するため、「開いて保存しただけ」の差分は検出されない。

### どれを使うか

- **デフォルトは slide-compare-lite**。「変更点を教えて」「差分を取って」など特に指定がない場合。
- **slide-compare** を使う場合:
  - lite で見えなかった変更を探したい場合（正規化で除去されている可能性）
  - 書式・属性レベルの変化まで漏れなく確認したい場合

## 2層連携フロー

```
Layer 1: slide-compare-lite
  ファイルメタ取得(I/Oのみ) → XML抽出(強正規化) → Jaccardマッピング → 正規化unified diff
  出力: before_meta / after_meta / summary / diff_dir / mapping
  → さらに詳細が必要なら Layer 2 へ

Layer 2: slide-compare
  XML抽出(最小正規化) → マッピング(強正規化版で実施) → ノードベースdiff
  出力: XMLノード単位の全変化
```

## Step 0: 実行環境のセットアップ

> **Python要件:** Python >= 3.10 が必要。ランチャーが自動で venv を作成し requirements を
> インストールする（初回のみ PyPI へのネットワークアクセスが必要）。

ランチャーに strategy（`lite` または `full`）と pptx パスを渡すだけでよい。venv 管理は不要。

**`<SKILL_DIR>` はこのスキルが配置されているディレクトリの絶対パスに読み替えること**
（例: `~/.claude/skills/pptx-compare`）。以降のドキュメントでも同じ表記を使う。

```bash
# macOS / Linux / Git Bash
<SKILL_DIR>/run.sh lite before.pptx after.pptx
<SKILL_DIR>/run.sh full before.pptx after.pptx
```

```powershell
# Windows PowerShell（bash 不要。引数は run.sh と同じ）
<SKILL_DIR>\run.ps1 lite before.pptx after.pptx
<SKILL_DIR>\run.ps1 full before.pptx after.pptx
```

## 使用方式の明示

比較を実行する前に、選択した方式と選択理由を以下の形式で必ず出力すること:

```
**使用方式:** <slide-compare-lite | slide-compare>
**選択理由:** <なぜその方式を選んだか（ユーザー指示・前層の結果・ケースの該当理由など）>
```

## 各方式の詳細

選択した方式のファイルを読み込み、そこに記載されたワークフローに従って実行する。

- **slide-compare-lite** → `<SKILL_DIR>/slide-compare-lite/slide-compare-lite.md`
- **slide-compare** → `<SKILL_DIR>/slide-compare/slide-compare.md`

## 注目すべき3要素

どの層でも、以下の3要素に着目して変更を報告する:

| 要素        | XMLタグ                           | 着目ポイント                                |
| ----------- | --------------------------------- | ------------------------------------------- |
| **textbox** | `<p:sp>`                          | テキスト内容 (`<a:t>`) の追加・削除・変更   |
| **table**   | `<p:graphicFrame>` (table URI)    | セル値、行・列の追加・削除                  |
| **chart**   | `<p:graphicFrame>` (chart URI)    | データ値 (`<c:v>`)、タイトル、凡例の変化    |
| **ノート**  | `<p:notes>` (diff 末尾に埋め込み) | ノートテキスト (`<a:t>`) の追加・削除・変更 |

## 注意事項

- PowerPoint はファイルを開いて保存した際にテーブルサイズを再計算することがある。
  テーブルサイズのみの変化は、ユーザーの編集ではなくこの自動再計算による可能性が高い。
- Jaccardマッピングの精度は高いため、スライドの追加・削除・並び替えも正確に検出できる。
