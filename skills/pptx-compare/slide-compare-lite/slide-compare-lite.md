---
name: slide-compare-lite
description: A high-speed difference detection tool for quickly identifying changes. It uses CRC32 pre-matching and strongly normalized Jaccard to identify changed slides, including slide masters, layouts, and themes. Use `slide-compare` if you need details on the changes.
---

# Slide Compare Lite

**「変更があったかどうか」を素早く把握するためのツール。**
CRC32チェックサムによる事前フィルタリングで未変更ファイルをスキップし、変更があったスライドだけJaccardマッピングを行う。
スライド本体に加え、スライドマスター・レイアウト・テーマも比較対象に含む。

変更の有無が確認できたら、詳細な変更内容の把握には `slide-compare` を使う。

## Usage

```bash
<SKILL_DIR>/run.sh lite before.pptx after.pptx
```

```powershell
# Windows PowerShell
<SKILL_DIR>\run.ps1 lite before.pptx after.pptx
```

## Options

| Option             | Default  | Description            |
| ------------------ | -------- | ---------------------- |
| `--threshold`      | `0.5`    | スライドマッチング閾値 |
| `--output-dir DIR` | temp dir | diffファイルの出力先   |

## Pipeline

1. **File Meta**: ファイルサイズ・更新日時をファイルシステムから取得（XML処理なし）
2. **Checksum**: ZIPエントリヘッダからCRC32を取得（全対象ファイル、XMLパースなし）
3. **Extract**: 強正規化（PowerPoint由来ノイズ除去 + テキストランマージ等）でスライド・マスター・レイアウト・テーマのXMLを読み込む
4. **Map**: Pass 1 = CRC一致 → `identical`（Jaccard不要）、Pass 2 = 残りをJaccardマッピング
5. **Diff**: 変更スライドのunified diff生成（`identical` はスキップ）
6. **Output**: メタ情報・サマリー・マッピング・diffパスをまとめて出力

## Output

```json
{
  "master_summary": {
    "identical": 3,
    "modified": 1,
    "added": 0,
    "deleted": 0
  },
  "master_mapping": [
    {
      "filename": "ppt/slideMasters/slideMaster1.xml",
      "before_index": 1,
      "after_index": 1,
      "status": "modified",
      "similarity": 0.0,
      "diff_file": "<diff_dir>/masters/master_B1_A1.diff",
      "before_checksum": "0x1a2b3c4d",
      "after_checksum": "0x5e6f7a8b"
    }
  ],
  "before_meta": {
    "path": "before.pptx",
    "size_bytes": 145000,
    "modified_at": "2026-01-15T10:00:00+09:00",
    "slide_count": 20
  },
  "after_meta": {
    "path": "after.pptx",
    "size_bytes": 160000,
    "modified_at": "2026-01-20T14:30:00+09:00",
    "slide_count": 21
  },
  "summary": {
    "identical": 17,
    "modified": 2,
    "added": 1,
    "deleted": 0
  },
  "diff_dir": "<一時ディレクトリ>/pptx_lite_compare_xxxxxxxx",
  "mapping": [
    {
      "before_index": 1,
      "after_index": 1,
      "status": "identical",
      "similarity": 1.0,
      "before_creation_id": "1000000001",
      "before_checksum": "0xaaaa0001",
      "after_creation_id": "1000000001",
      "after_checksum": "0xaaaa0001"
    },
    {
      "before_index": 10,
      "after_index": 10,
      "status": "modified",
      "similarity": 0.87,
      "diff_file": "<diff_dir>/slide_B10_A10.diff",
      "before_creation_id": "1000000010",
      "before_checksum": "0xbbbb0010",
      "after_creation_id": "1000000010",
      "after_checksum": "0xcccc0010"
    },
    {
      "before_index": null,
      "after_index": 21,
      "status": "added",
      "similarity": 0.0,
      "after_creation_id": "1000000021",
      "after_checksum": "0xdddd0021"
    }
  ]
}
```

## 実行後アクション

1. `summary` と `master_summary` で変更の有無を確認する
2. `before_meta` / `after_meta` でファイルサイズ・スライド枚数の変化を確認する
3. **変更がある場合slide-compareを実行し詳細を確認することを進める**

```bash
<SKILL_DIR>/run.sh full before.pptx after.pptx
# Windows PowerShell: <SKILL_DIR>\run.ps1 full before.pptx after.pptx
```

4. 変更がない場合（すべて `identical` または `added`/`deleted` のみ）は以下を報告する:

---

## レポートフォーマット（Lite）

変更の有無と規模のみを報告する。diff の詳細内容は記載しない。

```
## 比較結果（Lite）

### スライド
- Before: X枚 / After: Y枚
- 変更なし: N枚
- 変更あり: N枚（Before スライド A, B, ...）
- 追加: N枚 / 削除: N枚

### マスター・レイアウト・テーマ
- 変更なし: N件
- 変更あり: N件（ファイル名を列挙）
- 追加: N件 / 削除: N件

### 判定
変更が検出されました。詳細は `slide-compare` を実行して確認してください。
```

変更が一切ない場合:

```
### 判定
変更は検出されませんでした。
```
