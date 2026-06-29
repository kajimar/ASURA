# ASURA プロジェクト詳細説明

ASURA は、入力PDFからPowerPoint資料を量産するための「テンプレート + ルール + 検証」型のPPTX生成システムです。既存の80スライド規模のゴールドスタンダードPPTXデッキを分析し、そのデザイン・構造・制約を再利用可能なテンプレートとルールに落とし込むことで、別のPDFからもレイアウトを破綻させずにPPTXを生成することを目指しています。

このプロジェクトで最も重要な方針は、資料の見た目と根拠を同時に守ることです。内容をきれいに要約するだけではなく、数値や主張が入力PDFに裏付けられていること、各スライドのテキスト量がテンプレートの許容範囲に収まること、生成されたPPTXが後工程で壊れないことを検証対象にします。

## 目的

ASURA の目的は、PDF資料をもとにPPTXを機械的・再現可能に生成することです。単発の変換ツールではなく、ゴールドスタンダードPPTXから抽出したテンプレート、レイアウト部品、スロット制約、ルーティングルールを使って、同じ品質基準で多数の資料を作ることを前提にしています。

主なゴールは次の通りです。

- PDFからページ単位・チャンク単位の情報を抽出する。
- 抽出した情報をスライド構成に割り当てる。
- 入力特徴に応じて適切なスライドコンポーネントを選ぶ。
- 各コンポーネントのスロットに、制約内で文章・箇条書き・引用を流し込む。
- PPTXをレンダリングし、スキーマ・レイアウト・オーバーフローを検証する。
- 失敗した場合は修復ループを回し、必要に応じて人間の確認に回す。

## 全体アーキテクチャ

ASURA は、次の流れでPPTXを生成する設計です。

```text
parse PDF -> outline -> routing -> fill -> render -> validate -> overflow check
  -> repair loop (max N) -> optional human gate
```

現在の実装では、このうち `extract`、`blueprint`、`validate`、`render` の基本パスが CLI として用意されています。将来的なルーティングルール、フィルポリシー、修復ループ、人間による確認ゲートは、AGENTS.md に記載されたフェーズ設計に沿って拡張する想定です。

## 開発フェーズ

### Phase A: ゴールドPPTX分析

既存の80スライドのゴールドスタンダードPPTXを分析し、デザインと構造を機械的に再利用できる形式に分解します。

生成対象は次のような成果物です。

- `theme.json`: ページサイズ、フォント、色、フッター、引用スタイルなどのテーマ情報。
- `components/*.json`: 表紙、目次、本文、章区切り、引用一覧などのレイアウト部品。
- `slots`: `TITLE`、`BULLETS`、`ITEMS` など、内容を差し込む領域。
- `constraints`: 最大文字数、最大行数、最小フォントサイズ、縮小可否、オーバーフロー時の扱い。

このフェーズの成果は、`src/asura/templates/theme_default/template.json` のようなテンプレート定義として表現されます。

### Phase B: 入力PDF解析

入力PDFをページ単位のチャンクに分解し、テキスト、正規化済みテキスト、bbox、見出しレベル、数値などを抽出します。

標準の抽出器は `src/asura/core/extract/pdf_extractor.py` にあり、よりリッチなPDF抽出器として `src/asura/core/athra_pdf/` に Athra PDF Extractor が用意されています。

出力は `extraction.json` です。

### Phase C: スライドとチャンクの対応付け

抽出されたPDFチャンクを、生成するスライドに対応付けます。設計上は、自動候補マッチング、難しいケースでのLLM補助、最終的な人間確認ゲートを組み合わせる想定です。

出力は `alignment.json` で、概念的には `slide_id -> chunk_ids[]` の対応を持ちます。

### Phase D: ルール推論

ゴールドデッキと対応付け結果から、PDFの入力特徴をどのコンポーネントにルーティングするか、各スロットにどう書き込むか、デッキ全体をどう構成するかをルール化します。

想定される成果物は次の通りです。

- `routing_rules.yaml`: 入力特徴からコンポーネントを選ぶルール。
- `fill_policies.yaml`: スロットへの書き込み制約、要約・分割・引用の方針。
- `deck_policies.yaml`: 表紙、目次、章区切り、引用一覧などデッキ全体の構成方針。

### Phase E: テンプレート完成

テンプレート、ルール、検証が安定して動く状態を目指します。ここでは「一度だけ成功する」ことではなく、複数PDFに対して再現性をもってレイアウトが壊れないことが重要です。

## 主要データモデル

### `extraction.json`

PDFまたはPPTXから抽出した素材情報です。スキーマは `src/asura/core/schemas/extraction.schema.json` にあります。

主な構造は次の通りです。

- `schema_version`: 現在は `0.1`。
- `document`: `document_id`、`source_type`、`source_path`、`page_count` などの文書メタデータ。
- `chunks`: 抽出単位の配列。各チャンクは `chunk_id`、`text`、`normalized_text` を必須とし、必要に応じて `page`、`slide`、`bbox`、`bbox_pt`、`heading_level`、`numbers`、`style`、`image`、`table` などを持ちます。

現在の標準PDF抽出器では、PyMuPDFでテキストブロックを読み取り、ページ番号、bbox、テキスト、簡易見出しレベルをチャンクとして出力します。テキストを抽出できない画像PDFの場合は失敗します。

### `blueprint.json`

生成するスライドの設計図です。スキーマは `src/asura/core/schemas/blueprint.schema.json` にあります。

主な構造は次の通りです。

- `schema_version`: 現在は `0.1`。
- `document_id`: 元文書のID。
- `theme_id`: 使用するテーマID。
- `toc`: 目次項目。
- `slides`: スライド配列。

各スライドは、`component_id`、`slots`、`citations` を必須とします。`slots` には `TITLE` や `BULLETS` など、テンプレート側のスロット名に対応する値を入れます。`citations` は `※1` のような引用マーク、ページ番号、根拠チャンクIDを保持します。

現状の `blueprint` コマンドは v0.1 の決定的なプレースホルダー実装です。ページまたはスライド単位でチャンクをグループ化し、最初の非空行をタイトル、後続行を箇条書きとして `comp_title_bullets` に流します。

### `template.json`

PPTXのテーマとコンポーネント定義です。スキーマは `src/asura/core/schemas/template.schema.json` にあります。

デフォルトテンプレートは `src/asura/templates/theme_default/template.json` です。現在のデフォルトテーマは16:9、960 x 540 pt、`Noto Sans JP`、白背景、黒文字、青アクセントを基本にしています。

定義済みコンポーネントは次の通りです。

| component_id | role | 用途 |
| --- | --- | --- |
| `comp_cover` | `cover` | 表紙 |
| `comp_toc` | `toc` | 目次 |
| `comp_section_divider` | `section_divider` | 章区切り |
| `comp_title_bullets` | `body` | タイトル + 箇条書き本文 |
| `comp_citations` | `citations` | 引用一覧 |

各コンポーネントは `layout_elements` と `slots` を持ちます。スロットには `max_chars`、`max_lines`、`min_font_pt`、`shrink_to_fit`、`overflow_policy` が定義され、「内容よりもレイアウトを壊さない」ための制約として機能します。

### `runlog.json`

実行ログです。スキーマは `src/asura/core/schemas/runlog.schema.json` にあります。

`run_id`、入力情報、試行履歴、最終ステータスを保持します。試行履歴には `extract`、`blueprint`、`validate`、`render`、`postvalidate` といったステップ、ステータス、エラー、成果物パスが記録されます。最大試行回数はスキーマ上は3回までです。

## CLI

CLIエントリポイントは `uv run asura <command>` です。定義は `src/asura/apps/cli/main.py` にあります。

### `paths`

プロジェクトルートと各スキーマの場所を表示します。

```sh
uv run asura paths
```

### `check`

指定した run ディレクトリに、必要なインスタンスファイルとスキーマファイルが存在するか確認します。

```sh
uv run asura check --run runs/sample
```

期待されるインスタンスファイルは次の通りです。

- `template.json`
- `extraction.json`
- `blueprint.json`
- `runlog.json`

### `validate`

run ディレクトリ内のJSONを Draft 2020-12 JSON Schema で検証します。

```sh
uv run asura validate --run runs/sample
```

検証対象は `template`、`extraction`、`blueprint`、`runlog` です。

### `extract`

PDFまたはPPTXから `extraction.json` を生成します。

```sh
uv run asura extract input/source.pdf --out runs/my_run/extraction.json
uv run asura extract input/source.pptx --out runs/my_run/extraction.json --extended
```

PDFの場合は `asura.core.extract.pdf_extractor.extract_pdf`、PPTXの場合は `asura.core.extract.pptx_extractor.extract_pptx` が使われます。PPTXでは `--extended` を付けることで、スタイル、z-order、回転、画像、テーブルなどの拡張情報を含められます。

### `blueprint`

`extraction.json` から `blueprint.json` を生成します。

```sh
uv run asura blueprint --run runs/my_run
```

現在はプレースホルダー実装で、各ページの最初のテキストをタイトル、後続テキストを箇条書きとして扱います。生成後、`blueprint.schema.json` で検証してから書き込みます。

### `render`

`template.json`、`extraction.json`、`blueprint.json`、`runlog.json` を検証した上でPPTXを生成します。

```sh
uv run asura render --run runs/my_run --out runs/my_run/output.pptx
uv run asura render --run runs/my_run --out runs/my_run/output_dom.pptx --mode dom
```

`--mode template` は通常の生産パスで、テンプレートとブループリントのスロットからPPTXを構成します。`--mode dom` は拡張抽出フィールドからPPTXを再構成するモードで、既存PPTXのピクセル寄り再現に使います。

### `render-classified`

分類済みテンプレートJSONから、テンプレート仕様ディレクトリを参照してPPTXを直接生成します。

```sh
uv run asura render-classified runs/sample/kanji_deck_template_classified_v1.json \
  --out runs/out/output.pptx
```

このコマンドは、`input/templates_spec` とテンプレート抽出済みrun群を使って、分類済みページデータをDOMレンダリング用の抽出データに変換します。

## Athra PDF Extractor

Athra PDF Extractor は `src/asura/core/athra_pdf/` にある、標準PDF抽出器よりもリッチな抽出サブパッケージです。仕様は `docs/specs/athra_pdf_extractor.md` にまとまっています。

公開APIは次の通りです。

```python
from asura.core.athra_pdf import (
    extract_athra_pdf,
    render_debug_html,
    render_debug_png,
    run_contract_test,
    build_report,
    write_report,
)
```

Athra は次の特徴を持ちます。

- PyMuPDFの `page.get_text("dict")` を使って、テキスト選択可能なPDFを解析する。
- OCRや外部HTTP、LLM呼び出しは行わない。
- NFKC正規化、箇条書き記号除去、数値抽出を行う。
- フォントサイズ、太字、文字数、ページ位置、見出しパターンなど複数の信号から見出しレベルを推定する。
- 近接する本文ブロックをマージし、見出し境界で意味的チャンクを作る。
- 繰り返し出現するテキストをヘッダー・フッターとしてマークする。
- HTML/PNGのデバッグレンダリングと、契約テスト、メトリクスレポートを提供する。

Athra の重要な契約として、bbox は常に `[x0, y0, x1, y1]` の配列です。辞書形式ではありません。また、各チャンクは `chunk_id`、`block_type`、`page_no`、`order`、`bbox`、`text`、`normalized_text`、`heading_level`、`hash`、`meta` などを持つことが想定されています。

## レンダリングモード

ASURA には2種類のレンダリングモードがあります。

### template mode

通常の生産パスです。`blueprint.json` の `slides[*].component_id` でテンプレートのコンポーネントを選び、`slots` の値を各レイアウト要素に流し込みます。

このモードでは、テンプレート側のスロット制約が重要です。`max_chars` や `max_lines` を超える場合は、縮小、分割、切り詰め、失敗などのポリシーに従って処理する設計です。

### dom mode

既存PPTXなどから抽出した拡張フィールドをもとに、図形、テキスト、画像、テーブルを座標ベースで再描画するモードです。

ゴールドPPTXの分析や、テンプレート再現性の検証に向いています。`--extended` 付きのPPTX抽出と組み合わせることで、スタイル、透明度、回転、z-order、画像ペイロードなどを使った再構築が可能になります。

## ディレクトリ構成

```text
src/asura/
  apps/cli/
    main.py                     # asura CLI
  core/
    athra_pdf/                  # リッチPDF抽出器
    blueprint/                  # blueprint生成関連
    extract/                    # PDF/PPTX抽出器
    render/                     # PPTXレンダラー
    schemas/                    # JSON Schema
    utils/                      # スキーマ検証ユーティリティ
    validate/                   # 検証関連
  templates/
    theme_default/template.json # デフォルトテンプレート

docs/specs/
  athra_pdf_extractor.md        # Athra PDF Extractor仕様

input/
  PDFs, gold PPTX, template specs

runs/
  <run_id>/
    extraction.json
    blueprint.json
    output.pptx
    runlog.json
```

## 技術スタック

ASURA は Python 3.11 を前提とし、`uv` で実行・依存管理します。

主な依存ライブラリは次の通りです。

- `python-pptx`: PPTX生成・再構成。
- `pymupdf` (`fitz`): PDF解析、bbox取得、テキスト抽出。
- `jsonschema`: Draft 2020-12 JSON Schema によるデータ検証。

CLIスクリプトは `pyproject.toml` の `[project.scripts]` で `asura = "asura.apps.cli.main:main"` として定義されています。

## 制約と品質方針

ASURA では次の制約が強く優先されます。

- 入力に根拠のない数値や事実を生成しない。
- 数値・主張はPDFチャンクや引用に結び付ける。
- 「完璧な言い回し」よりも「レイアウトを壊さない」ことを優先する。
- テンプレートのスロット制約を守る。
- 修復ループの最大回数は設定可能にし、デフォルトは3回とする。
- 検証に失敗した状態でレンダリングを進めない。
- 画像PDFなど、テキスト抽出できないPDFは明示的に失敗させる。

この方針により、ASURA は生成AI的な自由作文ではなく、検証可能な資料生成パイプラインとして設計されています。

## 現状の実装上の注意

このリポジトリには、完成形の設計とv0.1の実装が混在しています。

現在動いている中心機能は、PDF/PPTX抽出、簡易blueprint生成、JSON Schema検証、PPTXレンダリングです。一方で、`alignment.json`、`routing_rules.yaml`、`fill_policies.yaml`、`deck_policies.yaml`、本格的な修復ループ、人間確認ゲートは設計上の重要要素ですが、すべてがCLIとして完成しているわけではありません。

また、標準 `extraction.schema.json` と Athra の契約にはフィールド名の違いがあります。標準スキーマでは `page` や辞書形式の `bbox` も扱いますが、Athra 仕様では `page_no` と配列形式の `bbox` を必須契約として扱います。Athra の出力を標準パイプラインに接続する場合は、この差分を明示的に吸収するアダプタまたはスキーマ整合が必要です。

## 典型的な作業例

新しいPDFからPPTXを生成する最小フローは次のようになります。

```sh
mkdir -p runs/my_run
cp src/asura/templates/theme_default/template.json runs/my_run/template.json

uv run asura extract input/source.pdf --out runs/my_run/extraction.json
uv run asura blueprint --run runs/my_run

# runlog.json は現状の validate/render で必須なので、スキーマに合うログを用意する
uv run asura validate --run runs/my_run
uv run asura render --run runs/my_run --out runs/my_run/output.pptx
```

PPTX再現やテンプレート分析寄りの作業では、PPTXを拡張抽出してDOMモードでレンダリングします。

```sh
uv run asura extract input/gold.pptx --out runs/gold/extraction.json --extended
uv run asura render --run runs/gold --out runs/gold/output_dom.pptx --mode dom
```

## 今後の拡張ポイント

ASURA を量産パイプラインとして安定させるには、次の領域が重要です。

- ゴールドPPTXからのコンポーネントクラスタリング精度を上げる。
- PDFチャンクとスライド意図のアラインメントを明示的な成果物にする。
- ルーティングルールとフィルポリシーをYAMLとして管理する。
- スロットごとのオーバーフロー検出と自動修復を強化する。
- 引用・根拠ポリシーを生成物全体で強制する。
- Athra 出力と標準 `extraction.json` スキーマの差分を整理する。
- レンダリング後のPPTXを画像化して、視覚的な破綻を検出する。

ASURA の本質は、PDFからPPTXを作る単純な変換ではありません。ゴールドデッキから学習した型を守り、根拠のある内容だけを、壊れないレイアウトに収めて大量生産するための資料生成基盤です。
