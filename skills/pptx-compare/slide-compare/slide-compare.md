---
name: slide-compare
description: Outputs detailed information about changes using a node-based diff of the minimally normalized XML. This applies not only to the slide content itself but also to the master, layout, and theme. Use this after verifying changes with `slide-compare-lite`.
---

# Slide Compare

**「どこがどう変わったか」を詳細に把握するためのツール。**
最小限の正規化のみ適用し、XMLノード単位のdiffを生成する。
スライド本体に加え、スライドマスター・レイアウト・テーマも比較対象に含む。

`slide-compare-lite` で変更が確認された後に実行する。

## Usage

```bash
<SKILL_DIR>/run.sh full before.pptx after.pptx
```

```powershell
# Windows PowerShell
<SKILL_DIR>\run.ps1 full before.pptx after.pptx
```

## Options

| Option             | Default  | Description            |
| ------------------ | -------- | ---------------------- |
| `--threshold`      | `0.5`    | スライドマッチング閾値 |
| `--output-dir DIR` | temp dir | diffファイルの出力先   |

## Pipeline

1. **Extract (diff用)**: 最小正規化（noisy regex除去 + チャートインライン化のみ）でスライドXML・マスターXMLを読み込む
2. **Extract (マッピング用)**: 強正規化版でも同じファイルを読み込む（Jaccardの精度を保つため）
3. **Map slides**: 強正規化版でJaccardマッピング（CRC比較なし、全スライドを対象）
4. **Map masters**: 強正規化版でファイル名一致マッピング（マスター・レイアウト・テーマはファイル名固定のためJaccard不要）
5. **Diff**: 最小正規化版XMLでノードベースdiff生成（スライド・マスター両方）

各ファイルを合計4回読み込む（スライド×2、マスター×2）。CRC比較は行わない。

## ノードベースdiff

XMLタグの追加・削除による行ずれを防ぐため、各XMLノードを1行にシリアライズしてからdiffを取る。

```
# 各行の形式:
  DEPTH  <tag attr="val">text

# 例:
<p:sld>
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:p>
            <a:r>
              <a:t>hello world
```

## 小さいXMLの場合

正規化後のXMLが200行未満の場合、diffを生成せずBefore/Afterの両方を出力ファイルに書き出す。
Claude が直接読んで差分を把握する。

```
=== BEFORE: before.pptx/ppt/slides/slide10.xml ===
<p:sld>
  ...
=== AFTER: after.pptx/ppt/slides/slide10.xml ===
<p:sld>
  ...
```

## Output

```json
{
  "mapping": [
    {
      "before_index": 10,
      "after_index": 10,
      "status": "modified",
      "similarity": 0.87,
      "diff_file": "<diff_dir>/slide_B10_A10.diff"
    }
  ],
  "master_mapping": [
    {
      "filename": "ppt/slideMasters/slideMaster1.xml",
      "before_index": 1,
      "after_index": 1,
      "status": "modified",
      "diff_file": "<diff_dir>/masters/master_B1_A1.diff"
    }
  ],
  "diff_dir": "<一時ディレクトリ>/pptx_full_compare_xxxxxxxx"
}
```

## 実行後アクション

1. `mapping` と `master_mapping` で変更対象ファイルを把握する
2. スライドのdiff一括読み込み:
   ```bash
   cat "$DIFF_DIR"/*.diff
   ```
3. マスター・レイアウト・テーマのdiff一括読み込み:
   ```bash
   cat "$DIFF_DIR/masters"/*.diff
   ```
   `$DIFF_DIR` は出力 JSON の `diff_dir` の値。個別に読むのは禁止。
   PowerShell の場合は `Get-Content "$DIFF_DIR\*.diff"` を使う。
4. **下記「レポートフォーマット」に従って**変更を日本語で解説する

## lite との違い

| 項目     | slide-compare-lite                  | slide-compare                 |
| -------- | ----------------------------------- | ----------------------------- |
| 目的     | 変更の有無を把握                    | 変更内容の詳細を把握          |
| 正規化   | 強正規化（PowerPoint由来ノイズ除去） | 最小正規化のみ                |
| CRC比較  | あり（未変更を高速スキップ）        | なし（全ファイル必ずdiff）    |
| diff形式 | unified diff                        | ノードベースdiff              |
| 抽出回数 | 1回（マッピング兼用）               | 4回（スライド×2、マスター×2） |

slide-compare では以下は除去 **されない**（lite では除去）:

- `<a:endParaRPr>` の変化
- `a16:rowId` / `a16:colId` の変化
- `<c:lang>` の変化
- `<a:defRPr>` の lang/altLang の変化
- `<c:spPr>` 内 `<a:solidFill>` の変化
- テキストランの分割・結合

これらが多数出現する場合は PowerPoint が保存時に行う自動変換の影響であることが多い。

---

## 無視すべき差分

- `a16:colId` / `a16:rowId` の追加のみ
- `dirty`, `err` 属性の変化
- `fillRect` のゼロ値属性の省略

---

## 注意すべき差分

- テーブルサイズ**のみ**の変化（PowerPointによるファイル保存時の自動再計算が行われる可能性があるため）
  - 以下のような注釈を出力に含めることを推奨する。
  - ⚠️ PowerPoint はファイル保存時にテーブルサイズを自動再計算することがあります。以下のスライドでサイズのみの変化が検出されました。実際に変化しているかどうかは PowerPoint で開いて確認してください。

---

## レポートフォーマット

diff を読んで以下のフォーマットで報告すること。**類似度の数値は出力しない。**

---

### テンプレート

```
## 比較結果

### サマリー
- スライド数: Before X枚 / After Y枚
- 変更なし: N枚
- 変更あり: N枚（Before スライド A, B, ...）
- 追加: N枚 / 削除: N枚

### マスター・レイアウト・テーマ   ← 変更がなければ省略
- 変更あり: N件（ファイル名を列挙）

### スライドの並び替え   ← 並び替えがなければ省略

| Before スライド | After スライド | 移動 |
|---|---|---|
| 11 | 13 | +2 |
| 12 | 11 | -1 |
| 13 | 12 | -1 |
| （なし） | 68 | 新規追加 |

---

### Before スライドX / After スライドY   ← テーブルサイズのみの変化は除く

| 種別 | 変化の内容 |
|---|---|
| テキスト | 「旧テキスト」→「新テキスト」 |
| テキスト（書式） | 太字・14pt・白色 → デフォルト書式 |
| テーブル 2行3列 | 「旧セル値」→「新セル値」 |
| テーブル（位置） | 左に移動 |
| チャート（系列「売上」） | 3番目: 1,500 → 1,800 |
| チャート（左Y軸） | 非表示になりました |
| チャート（タイトル） | 「旧タイトル」→「新タイトル」 |
| ノート | 「旧ノートテキスト」→「新ノートテキスト」 |

---
（次のスライドへ）

---

### マスター・レイアウト・テーマの変更   ← 変更がなければ省略

| ファイル | 変化の内容 |
|---|---|
| slideMaster1.xml | フォント: Meiryo → Yu Gothic |
| slideLayout3.xml | 背景色: #FFFFFF → #F0F0F0 |

---

### テーブルサイズの自動再計算   ← 2スライド以上の場合はまとめて出力、1スライドのみなら上の表に含める

⚠️ PowerPoint はファイル保存時にテーブルサイズを自動再計算することがあります。以下のスライドでサイズのみの変化が検出されました。実際に変化しているかどうかは PowerPoint で開いて確認してください。

| Before スライド | After スライド | 変化の内容 |
|---|---|---|
| 11 | 13 | 高さ: 893115 → 474795 EMU |
| 19 | 21 | 幅: 11849987 → 11849991 EMU（わずかな変化） |

---

### まとめ
（全体の変更点を2〜3行で要約）
```

---

### 種別の書き方ルール

**テキスト系:**

| 変化         | 種別欄             | 変化の内容欄                                          |
| ------------ | ------------------ | ----------------------------------------------------- |
| テキスト変更 | `テキスト`         | 「旧」→「新」                                         |
| テキスト削除 | `テキスト`         | 「削除されたテキスト」が削除                          |
| テキスト追加 | `テキスト`         | 「追加されたテキスト」が追加                          |
| 書式変更     | `テキスト（書式）` | 変更内容を言葉で（例: 太字・14pt・白色 → デフォルト） |

**テーブル系:**

diff 中の `<a:tr>` の出現順（1始まり）を行番号、`<a:tc>` の出現順（1始まり）を列番号とする。

| 変化               | 種別欄               | 変化の内容欄                                   |
| ------------------ | -------------------- |------------------------------------------|
| セル値変更         | `テーブル N行M列`    | 「旧」→「新」                                  |
| 位置変更           | `テーブル（位置）`   | 方向を言葉で（例: 左に移動）                          |
| サイズ変更（単独） | `テーブル（サイズ）` | 高さ/幅: 旧値 → 新値 EMU ⚠️ PowerPoint自動再計算の可能性 |

**チャート系:**

XMLタグをそのまま書かず意味を理解できる言葉で記載すること。

| diff で見えるパターン  | 種別欄                     | 変化の内容欄                     |
| ---------------------- | -------------------------- | -------------------------------- |
| `<c:v>` の値変化       | `チャート（系列「名前」）` | N番目: 旧値 → 新値               |
| `<c:delete val="1">`   | `チャート（X軸）` など     | 非表示になりました               |
| `<c:delete val="0">`   | `チャート（X軸）` など     | 表示されるようになりました       |
| `<c:spPr>` 削除（軸）  | `チャート（X軸）` など     | スタイルがデフォルトに戻りました |
| `off x` / `off y` 変化 | `チャート（位置）`         | 右に移動、など方向を言葉で       |
| `<c:title>` 変化       | `チャート（タイトル）`     | 「旧」→「新」                    |
| `<c:ser>` 追加/削除    | `チャート（系列）`         | 「名前」が追加/削除されました    |

軸の特定: `<c:valAx>` → Y軸、`<c:catAx>` → X軸、`axPos val="l/r"` → 左/右軸、`axPos val="b/t"` → 下/上軸

**ノート系:**

diff 末尾の `<p:notes>` ブロック内の変化がノートの変更。

| 変化         | 種別欄   | 変化の内容欄                 |
| ------------ | -------- | ---------------------------- |
| テキスト変更 | `ノート` | 「旧」→「新」                |
| テキスト追加 | `ノート` | 「追加されたテキスト」が追加 |
| テキスト削除 | `ノート` | 「削除されたテキスト」が削除 |
