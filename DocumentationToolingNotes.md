# <a name="how-the-office-javascript-api-documentation-is-generated"></a>Office JavaScript API ドキュメントの生成方法

Office JavaScript リファレンスドキュメントページは、型定義ファイルおよびスニペットの例から生成されます。 このプロセスでは、オープンソースツールとリポジトリ固有のスクリプトの組み合わせを使用します。 このドキュメントは、このリポジトリのプロセスを透過的にすることを目的としており、コミュニティがこのコンテンツをより効果的に活用し、投稿できるようにします。

## <a name="content-sources"></a>コンテンツ ソース

Office JS リファレンスドキュメントを作成するために、次の2種類のコンテンツが組み合わされています。型定義とコードスニペット。 これにより、完全な API の適用範囲が保証され、小さなインラインコードサンプルが提供されます。

### <a name="type-definition-files"></a>型定義ファイル

完全に[入力されている](https://github.com/DefinitelyTyped/DefinitelyTyped)型定義ファイルは、ドキュメントの真実の単一のソースです。 TypeScript を使用する Office アドインは、これらの型定義ファイルを使用してコンパイルされます。 これらは、JavaScript および TypeScript 開発者にも IntelliSense 機能を提供します。 これらの定義から参照ドキュメントを作成することで、より正確な情報が得られます。

各ドキュメントのサブセクションのソースコンテンツを提供する、4つの関連する d. ファイルがあります。

- [office-js/d-u-n-s](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts) (リリースの定義)
  - [Excel (リリース)](https://docs.microsoft.com/javascript/api/excel_release)
  - [OneNote](https://docs.microsoft.com/javascript/api/onenote)
  - [PowerPoint](https://docs.microsoft.com/javascript/api/powerpoint)
  - [Visio](https://docs.microsoft.com/javascript/api/visio)
  - [Word (リリース)](https://docs.microsoft.com/javascript/api/word_release)
  - [共通 API の OfficeExtensions サブセクション](https://docs.microsoft.com/javascript/api/office)
- [office-js-プレビュー/インデックス](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)(プレビューの定義)
  - [Excel (プレビュー)](https://docs.microsoft.com/javascript/api/excel)
  - [Outlook (プレビュー)](https://docs.microsoft.com/javascript/api/outlook)
  - [Word (プレビュー)](https://docs.microsoft.com/javascript/api/word)
  - [共通 API](https://docs.microsoft.com/javascript/api/office)
- [カスタム-関数-ランタイム/インデックス](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/custom-functions-runtime/index.d.ts)(Excel カスタム関数ランタイム定義)。
  - [カスタム関数](https://docs.microsoft.com/javascript/api/custom-functions-runtime)
- [office ランタイム/インデックス](https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/office-runtime/index.d.ts)(カスタム関数プラットフォーム用の office ランタイム定義) を使用します。
  - [Office ランタイム](https://docs.microsoft.com/javascript/api/office-runtime)

以前のバージョンの Api には独自の d-u-n-s ファイルがあります。 これらは、新しい API 要件セットがリリースされるときに保持されます。 また、[バージョン Remover ツール](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/tools/VersionRemover.ts)を使用して生成することもできます。 これらの古い d-u-n-s ファイルは保持されているため、イベント Api にパッチまたは変更があっても、元の動作はまだ文書化されています。 これは、古いバージョンの API を対象にする必要がある場合に役立ちます。

#### <a name="testing-type-definition-file-changes"></a>種類の定義ファイルの変更をテストする

Office JavaScript API のドキュメント変更は、前述の4つの d-u-n-s ファイルを編集することによって行われます。 ただし、markdown を作成する前に、(たとえば、書式設定がにどのように変換されるかをテストする必要がある場合は)、「[ドキュメントの生成](https://github.com/OfficeDev/office-js-docs-reference/tree/master/generate-docs/script-inputs)」および「 [generatedocs](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd)」で対応するファイルを編集することによって、PR を入力する前に、変更をテストすることができます。 メッセージが表示されたら、[ローカルファイル] オプションを選択します。

このリポジトリのリモートブランチに変更を加えると、docs.microsoft.com プラットフォームによってテストブランチが構築されます。 この分岐は review.docs.microsoft.com に表示されます。これは、Microsoft の内部担当者のみがアクセスできます。 PR をレビューしているユーザーは、レビューサイトが正確であるかどうかを確認します。

### <a name="code-snippets"></a>コード スニペット

コード例スニペットは、2つのソースから参照ページに追加されます。

- [スクリプトラボサンプル](https://github.com/OfficeDev/office-js-snippets)
- [ローカルコードスニペット](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/code-snippets)

ローカルスニペットは、ホスト固有の yaml ファイルにあります。 これらのコンテンツはクラスとフィールドによって整理されているため、参照ページ内の適切な場所にマップできます。 スニペットの言語 (JavaScript または TypeScript) は、await ステートメントの使用によって推論されます。

スクリプトラボスニペットは、作業用サンプルから引き出されています。 現時点では、Excel、Outlook、および Word のサンプルは、[マッピングファイル](https://github.com/OfficeDev/office-js-snippets/tree/master/snippet-extractor-metadata)によって参照ドキュメントのセクションにマッピングされています。 これらは、個々のサンプルメソッドを API のプロパティまたはメソッドと照合します。 Office js-スニペットリポジトリを`yarn start`実行すると、すべてのマップされたスニペットを含む[yaml ファイル](https://github.com/OfficeDev/office-js-snippets/blob/master/snippet-extractor-output/snippets.yaml)が作成されます。 この yaml ファイルは、リファレンスドキュメントツールへの入力です。

## <a name="tooling-pipeline"></a>ツールパイプライン

![明示的に入力された制御フロー、プリプロセッサ、API 抽出器、midprocessor、API の解析、およびポストプロセッサへの制御フローを示すイメージ。](ToolingPipeline.png)

コンテンツソースと最終ページの間では、ドキュメントのコンテンツは5つのツールステップを通過します。

1. [プリプロセッサスクリプト](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/preprocessor.ts)
1. [API 抽出器](https://api-extractor.com/)
1. [Midprocessor スクリプト](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/midprocessor.ts)
1. [API の解析の方法](https://github.com/microsoft/rushstack/blob/master/apps/api-documenter/README.md)
1. [ポストプロセッサスクリプト](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/scripts/postprocessor.ts)

プリプロセッサは、d-u-n-s ファイルを取得し、それらをホスト固有のセクションに分割します。 今後のツールがデータを適切に処理するために必要なクリーンアップが実行されます。

API エクストラクターは、d-u-n-s ファイルを JSON データに変換します。 このトークンは、解析を簡単にするために、すべての型データを有効にします。

Midprocessor は、コードスニペットを取得し、それらを適切なホストとペアにして、Outlook と共通 API オブジェクト間のクロスリンクをクリーンアップします。

API 解析ツールは、JSON データを yml ファイルに変換します。 Yml ファイルは、ドキュメントを docs.microsoft.com に公開する Open Publishing システムによって markdown に変換されます。 API の解析には、コードスニペットを挿入する Office 固有の拡張機能も含まれています。

ポストプロセッサは、目次をクリーンアップし、yml ファイルを[発行フォルダー](https://github.com/OfficeDev/office-js-docs-reference/tree/master/docs/docs-ref-autogen)に移動します。

次の5つの手順はすべて、 [Generatedocs](https://github.com/OfficeDev/office-js-docs-reference/blob/master/generate-docs/GenerateDocs.cmd)が実行されたときに実行されます。 このスクリプトは、ノードモジュールのインストール、古いファイルセットのクリーンアップ、および各要件セットのバージョンタイプの定義ファイルも処理します。
