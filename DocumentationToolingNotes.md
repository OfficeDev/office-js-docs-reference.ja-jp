# <a name="how-the-office-javascript-api-documentation-is-generated"></a>JavaScript API Office生成される方法

JavaScript Officeドキュメント ページは、型定義ファイルとサンプル スニペットから生成されます。 このプロセスでは、オープン ソース ツールとリポジトリ固有のスクリプトのブレンドを使用します。 このドキュメントでは、コミュニティがこのコンテンツのメリットを高め、投稿できるよう、このリポジトリのプロセスを透過的に提供します。

## <a name="content-sources"></a>コンテンツ ソース

Office-JS リファレンス ドキュメントを作成するために、型定義とコード スニペットの 2 種類のコンテンツが組み合わされます。 これらは完全な API カバレッジを確保し、小規模なインライン コード サンプルを提供します。

### <a name="type-definition-files"></a>型定義ファイル

[「](https://github.com/DefinitelyTyped/DefinitelyTyped)間違いなく型指定」の型定義ファイルは、ドキュメントの単一の真の情報源です。 TypeScript Office使用するアドインは、これらの型定義ファイルを使用してコンパイルされます。 また、JavaScript と TypeScript の開発者は、IntelliSenseできます。 これらの定義から参照ドキュメントを作成することで、より正確な情報が提供されます。

ドキュメントの異なるサブセクションのソース コンテンツを提供する 4 つの関連する d.ts ファイルがあります。

- [office-js/index.d.ts](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts) (リリース定義)
  - [Excel (リリース)](https://docs.microsoft.com/javascript/api/excel_release)
  - [OneNote](https://docs.microsoft.com/javascript/api/onenote)
  - [PowerPoint](https://docs.microsoft.com/javascript/api/powerpoint)
  - [Visio](https://docs.microsoft.com/javascript/api/visio)
  - [Word (リリース)](https://docs.microsoft.com/javascript/api/word_release)
  - [共通 API の OfficeExtensions サブセクション](https://docs.microsoft.com/javascript/api/office)
- [office-js-preview/index.d.ts](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts) (プレビュー定義)
  - [Excel (プレビュー)](https://docs.microsoft.com/javascript/api/excel)
  - [Outlook (プレビュー)](https://docs.microsoft.com/javascript/api/outlook)
  - [Word (プレビュー)](https://docs.microsoft.com/javascript/api/word)
  - [共通 API](https://docs.microsoft.com/javascript/api/office)
- [custom-functions-runtime/index.d.ts](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/custom-functions-runtime/index.d.ts) (Excel カスタム関数ランタイム定義)
  - [カスタム関数](https://docs.microsoft.com/javascript/api/custom-functions-runtime)
- [office-runtime/index.d.ts](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-runtime/index.d.ts) (カスタム関数プラットフォームの Office ランタイム定義)。
  - [Office ランタイム](https://docs.microsoft.com/javascript/api/office-runtime)

以前のバージョンの API には独自の d.ts ファイルがあります。 これらは、新しい API 要件セットがリリースされると保持されます。 また、バージョン削除ツールを [使用して生成することもできます](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/tools/VersionRemover.ts)。 これらの古い d.ts ファイルは維持され、API に修正プログラムが適用または変更された場合でも、元の動作は文書化されます。 これは、API の古いバージョンをターゲットとする必要がある場合に便利です。

#### <a name="testing-type-definition-file-changes"></a>型定義ファイルの変更のテスト

JavaScript API のドキュメントOffice変更は、上記の 4 つの d.ts ファイルを編集することで行われます。 ただし [、generate-docs/script-inputs](https://github.com/OfficeDev/office-js-docs-reference/tree/master/generate-docs/script-inputs) で対応するファイルを編集し [、GenerateDocs.cmd](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd)を実行することで、PR を DefinitelyTyped に提出する前に変更をテストできます (たとえば、書式設定がマークダウンに変換される方法をテストする必要がある場合)。 メッセージが表示されたら、[ローカル ファイル] オプションを選択します。

このレポのリモート分岐に変更をプッシュすると、docs.microsoft.comプラットフォームでテスト分岐が作成されます。 この分岐は、内部の Microsoft review.docs.microsoft.comによってのみアクセス可能な場所にレンダリングされます。 PR を確認するユーザーは、レビュー サイトの正確性を確認します。

### <a name="code-snippets"></a>コード スニペット

コード例スニペットは、次の 2 つのソースから参照ページに追加されます。

- [Script Lab Samples](https://github.com/OfficeDev/office-js-snippets)
- [ローカル コード スニペット](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/code-snippets)

ローカル スニペットは、ホスト固有の yaml ファイルに含めます。 コンテンツはクラスとフィールドごとに整理され、参照ページ内の適切な場所にマップできます。 スニペットの言語 (JavaScript または TypeScript) は、await ステートメントの使用によって推測されます。

Script Lab スニペットは、動作するサンプルから取得されます。 現在、Excel、Outlook、PowerPoint、および Word のサンプルは、マッピング ファイルを使用してドキュメント セクションを参照するために [マップされています](https://github.com/OfficeDev/office-js-snippets/tree/prod/snippet-extractor-metadata)。 これらは、個々のサンプル メソッドを API のプロパティまたはメソッドに一致します。 office-js-snippets リポジトリが実行されると、マップされたスニペットを含む `yarn start` [yaml](https://github.com/OfficeDev/office-js-snippets/blob/prod/snippet-extractor-output/snippets.yaml) ファイルが作成されます。 この yaml ファイルは、リファレンス ドキュメント ツールへの入力です。

## <a name="tooling-pipeline"></a>ツール パイプライン

![間違いなく型指定されたコントロール フロー、プリプロセッサ、API 抽出器、中間プロセッサ、API Documenter、およびポストプロセッサへの制御フローを示す画像。](ToolingPipeline.png)

コンテンツ ソースと最終ページの間で、ドキュメント コンテンツは 5 つのツールステップを実行します。

1. [プリプロセッサ スクリプト](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/preprocessor.ts)
1. [API 抽出器](https://api-extractor.com/)
1. [Midprocessor スクリプト](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/midprocessor.ts)
1. [API Documenter](https://github.com/microsoft/rushstack/blob/master/apps/api-documenter/README.md)
1. [ポストプロセッサ スクリプト](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/postprocessor.ts)

プリプロセッサは d.ts ファイルを取得し、ホスト固有のセクションに分割します。 後続のツールがデータを適切に処理するために必要なクリーンアップを実行します。

API 抽出器は、d.ts ファイルを JSON データに変換します。 これにより、すべての型データがトークン化され、解析が容易になります。

Midprocessor は、コード スニペットを取得し、それらを適切なホストとペアにし、Outlook と共通 API オブジェクトの間の問題をクリーンアップします。

API Documenter は JSON データを .yml ファイルに変換します。 .yml ファイルは、ドキュメントを公開する Open Publishing System によってマークダウンdocs.microsoft.com。 API Documenter には、コード スニペットをOffice固有の拡張機能も含めます。

ポストプロセッサは目次をクリーンアップし、.yml ファイルを発行フォルダー [に移動します](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen)。

これらの 5 つの手順はすべて [、GenerateDocs.cmd を実行すると](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd) 実行されます。 このスクリプトは、ノード モジュールのインストールの処理、古いファイル セットの削除、および各要件セットの型定義ファイルのバージョンも処理します。
