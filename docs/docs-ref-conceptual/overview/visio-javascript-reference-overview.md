# <a name="visio-javascript-api-overview"></a>Visio の JavaScript API の概要

Visio JavaScript API を使うと、SharePoint Online で Visio の図を埋め込むことができます。 埋め込んだ Visio の図は、SharePoint ドキュメント ライブラリに保存され、SharePoint ページに表示されます。 Html に表示する Visio 図面を埋め込むには、`<iframe>`要素です。 そうすると、Visio JavaScript API を使用して、プログラムで埋め込み済みの図を使った作業ができるようになります。

![SharePoint ページの iframe 上にある Visio の図とスクリプト エディター Web パーツ](/javascript/api/docs-ref-conceptual/images/visio-api-block-diagram.png)


Visio JavaScript API を使用して、次のことを行えます。

* ページや図形のように Visio の図の要素と対話します。
* Visio の図のキャンバス上には、マークアップを作成します。
* 図面内でのマウス イベントのカスタム ハンドラーを記述します。
* 図形テキスト、図形データ、およびハイパーリンクなどの図のデータをソリューションに公開する。

この記事では、Visio Online で Visio JavaScript API を使って SharePoint Online のソリューションをビルドする方法について説明します。また、**EmbeddedSession**、**RequestContext**、JavaScript プロキシ オブジェクトなどの API、および **sync()**、**Visio.run()**、**load()** のメソッドを使用するために知っておくべき主な概念について紹介します。コード例により、これらの概念を適用する方法を示します。

## <a name="embeddedsession"></a>EmbeddedSession

EmbeddedSession オブジェクトは、開発者のフレームと Visio Online のフレームの間の通信を初期化します。

```js
var session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
session.init().then(function () {
    window.console.log("Session successfully initialized");
});
```

## <a name="visiorunsession-functioncontext--batch-"></a>Visio.run (セッション、function(context) {バッチ})

**Visio.run()** は、Visio オブジェクト モデルに対してアクションを実行するバッチ スクリプトを実行します。 このバッチ コマンドには、JavaScript のローカル プロキシ オブジェクトの定義と、ローカル オブジェクトと Visio オブジェクトの間で状態を同期し、解決される約束を返す **sync()** メソッドが含まれます。 **Visio.run()** で要求をバッチ処理する利点は、約束が解決されるときに、実行中に割り当てられたすべての追跡ページ オブジェクトが自動的に解放されることです。

メソッドの実行を選択し、セッションと RequestContext オブジェクトでは、約束を返します (通常は、ジャスト**context.sync()** の結果)。 バッチ操作は **Visio.run()** の外部で実行することができます。 ただし、このようなシナリオでは、ページ オブジェクトの参照は、手動で追跡および管理する必要があります。

## <a name="requestcontext"></a>RequestContext

RequestContext オブジェクトには、Visio アプリケーションへの要求が容易になります。 現像フレームと Visio のオンライン アプリケーションは、異なる 2 つの iframe で実行するため、開発者のフレームから Visio とページや図形などの関連するオブジェクトへのアクセスを取得する RequestContext オブジェクト (次の例の内容を含む) が必要です。

```js
function hideToolbars() {
    Visio.run(session, function(context){
        var app = context.document.application;
        app.showToolbars = false;
        return context.sync().then(function () {
            window.console.log("Toolbars Hidden");
        });
    }).catch(function(error)
    {
        window.console.log("Error: " + error);
    });
};
```

## <a name="proxy-objects"></a>プロキシ オブジェクト

アドインで宣言され使用される Visio の JavaScript オブジェクトは、Visio 図面の実際のオブジェクトのプロキシ オブジェクトになります。プロキシ オブジェクトで実行されたすべてのアクションは、Visio では認識されません。また、Visio ドキュメントの状態は、ドキュメントの状態が同期されるまでプロキシ オブジェクトで認識されません。ドキュメントの状態は、`context.sync()` の実行時に同期されます。

たとえば、ローカルの JavaScript オブジェクトの getActivePage は、選択したページを参照する宣言されています。 これは、このオブジェクトのプロパティと呼び出しメソッドの設定をキューに入れるために使用できます。 **Sync()** メソッドが実行されるまでこのようなオブジェクトのアクションはありません実現しています。

```js
var activePage = context.document.getActivePage();
```

## <a name="sync"></a>sync()

**Sync()** メソッドは、JavaScript プロキシ オブジェクト間の状態を同期コンテキスト上のキューに入れられた命令を実行して、Visio での実際のオブジェクトとのプロパティを取得中に、コードで使用する Office オブジェクトが読み込まれます。 このメソッドは、同期処理が完了したときに解決される約束を返します。 

## <a name="load"></a>load()

**load()** メソッドは、アドインの JavaScript レイヤーで作成されたプロキシ オブジェクトに設定を取り込むために使用されます。ドキュメントなどのオブジェクトを取得しようとすると、まず JavaScript レイヤーでローカル プロキシ オブジェクトが作成されます。このようなオブジェクトは、そのプロパティと呼び出しメソッドの設定をキューに登録するために使用できます。しかし、オブジェクトのプロパティや関係を読み取るためには、最初に **load()** メソッドと **sync()** メソッドを呼び出す必要があります。load() メソッドは、**sync()** メソッドが呼び出されたときに読み込まれる必要があるプロパティと関係を取り込みます。

以下に示すのは **load()** メソッドの構文です。

```js
object.load(string: properties); //or object.load(array: properties); //or object.load({loadOption});
```

1. **プロパティ**は、読み込まれていると指定されたコンマ区切りの文字列を指定するプロパティ名のリストまたは名前の配列です。 詳細については、各オブジェクトの下の **.load()** メソッドを参照してください。

2. **loadOption** は、selection、expansion、top、skip の各オプションについて説明するオブジェクトを指定します。詳細については、オブジェクトの読み込みの[オプション](/javascript/api/office/officeextension.loadoption)を参照してください。

## <a name="example-printing-all-shapes-text-in-active-page"></a>例:アクティブ ページですべての図形テキストを印刷する

次の例では、図形の配列オブジェクトから図形テキストの値を印刷する方法を示します。
**Visio.run()** メソッドには、命令のバッチが含まれています。 このバッチの一部として、作業中のドキュメントの図形を参照するプロキシ オブジェクトが作成されます。

これらすべてのコマンドはキューに登録し、 **context.sync()** が呼び出されたときに実行します。 **sync()** メソッドが返す約束は、このメソッドを他の操作とチェーンにするために使用できます。

```js
Visio.run(session, function (context) {
    var page = context.document.getActivePage();
    var shapes = page.shapes;
    shapes.load();
    return context.sync().then(function () {
        for(var i=0; i<shapes.items.length;i++) {
            var shape = shapes.items[i];
            window.console.log("Shape Text: " + shape.text );
        }
    });
}).catch(function(error) {
    window.console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        window.console.log ("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="error-messages"></a>エラー メッセージ

エラーは、コードとメッセージで構成される error オブジェクトを使用して返されます。次の表は、発生する可能性があるエラー状態の一覧を示しています。

| error.code            | error.message |
|-----------------------|----------------------------------------------------------------|
| InvalidArgument       | 引数が無効であるか、存在しません。または形式が正しくありません。 |
| GeneralException      | 要求の処理中に内部エラーが発生しました。 |
| NotImplemented        | 要求された機能は実装されていません。  |
| UnsupportedOperation  | 試行中の操作はサポートされていません。 |
| AccessDenied          | 要求された操作を実行できません。 |
| ItemNotFound          | 要求されたリソースは存在しません。 |

## <a name="get-started"></a>作業の開始

このセクションの例を使用するにを開始します。 この例では、プログラムを使用して Visio の図で選択した図形の図形のテキストを表示する方法を示します。 最初に、SharePoint Online でクラシックのページを作成または既存のページを編集します。 スクリプト エディターの web パーツをページに追加し、次のコードをコピーおよび貼り付けです。

```js
<script src='https://appsforoffice.microsoft.com/embedded/1.0/visio-web-embedded.js' type='text/javascript'></script>

Enter Visio File Url:<br/>
<script language="javascript">
document.write("<input type='text' id='fileUrl' size='120'/>");
document.write("<input type='button' value='InitEmbeddedFrame' onclick='initEmbeddedFrame()' />");
document.write("<br />");
document.write("<input type='button' value='SelectedShapeText' onclick='getSelectedShapeText()' />");
document.write("<textarea id='ResultOutput' style='width:350px;height:60px'> </textarea>");
document.write("<div id='iframeHost' />");

let session; // Global variable to store the session and pass it afterwards in Visio.run()
var textArea;
// Loads the Visio application and Initializes communication between developer frame and Visio online frame
function initEmbeddedFrame() {
    textArea = document.getElementById('ResultOutput');
    var url = document.getElementById('fileUrl').value;
    if (!url) {
        window.alert("File URL should not be empty");
    }
    // APIs are enabled for EmbedView action only.
    url = url.replace("action=view","action=embedview");
    url = url.replace("action=interactivepreview","action=embedview");
    url = url.replace("action=default","action=embedview");
    url = url.replace("action=edit","action=embedview");
  
    session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
    return session.init().then(function () {
        // Initialization is successful
        textArea.value  = "Initialization is successful";
    });
}

// Code for getting selected Shape Text using the shapes collection object
function getSelectedShapeText() {
    Visio.run(session, function (context) {
        var page = context.document.getActivePage();
        var shapes = page.shapes;
        shapes.load();
        return context.sync().then(function () {
            textArea.value = "Please select a Shape in the Diagram";
            for(var i=0; i<shapes.items.length;i++) {
                var shape = shapes.items[i];
                if ( shape.select == true) {
                    textArea.value = shape.text;
                    return;
                }
            }
        });
    }).catch(function(error) {
        textArea.value = "Error: ";
        if (error instanceof OfficeExtension.Error) {
            textArea.value += "Debug info: " + JSON.stringify(error.debugInfo);
        }
    });
}
</script>
```

その後、必要なものは使用する Visio の図の URL です。 だけで Visio 図面を SharePoint Online にアップロードし、オンラインの Visio で開くことです。 埋め込みダイアログ ボックスを開くし、埋め込みの URL を使用して、上の例で。

![埋め込むダイアログから Visio ファイルの URL をコピーします。](/javascript/api/docs-ref-conceptual/images/Visio-embed-url.png)

を使用している Visio のオンライン編集モードの場合は、埋め込むダイアログを開く**ファイル**を選択することによって > **共有** > **埋め込む**。 を使用している Visio のオンライン表示モードの場合は、']' と、**埋め込み**を選択することで埋め込み] ダイアログを開きます。

## <a name="open-api-specifications"></a>Open API の仕様

新しい API の設計と開発にあたり、[Open API の仕様](../openspec.md)ページでこれらに対するフィードバックの提供が可能になります。パイプラインの新機能をご確認いただき、設計の仕様に関する情報をお寄せください。

## <a name="visio-javascript-api-reference"></a>Visio の JavaScript API リファレンス

Visio の JavaScript API の詳細については、 [Visio の JavaScript API リファレンス ドキュメント](/javascript/api/visio)を参照してください。
