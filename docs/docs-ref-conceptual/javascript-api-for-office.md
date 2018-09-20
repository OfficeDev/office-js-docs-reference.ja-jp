# <a name="javascript-api-for-office"></a>JavaScript API for Office

Office 用の JavaScript API を使用すると、Office ホスト アプリケーションのオブジェクト モデルと対話する web アプリケーションを作成できます。 アプリケーションは、スクリプト ローダーは、office.js ライブラリを参照します。 Office.js ライブラリでは、アドインを実行している Office アプリケーションに適用可能なオブジェクト モデルを読み込みます。 次の JavaScript オブジェクト モデルを使用することができます。

- **一般的な Api**の**Office 2013**で導入された Api です。 これは**すべての Office ホスト アプリケーション**に読み込まれている、Office クライアント アプリケーションと、アドインのアプリケーションを接続します。 オブジェクト モデルには、Office クライアントに固有の Api と複数の Office クライアントのホスト アプリケーションに適用可能な Api が含まれています。 すべてのコンテンツは、 **API の共有です**。 

  **Outlook**では、API の共通の構文も使用します。 [エイリアスの Office のすべてのものには、Office ドキュメント、ワークシート、プレゼンテーション、メール アイテム、および Office アドインからプロジェクト内のコンテンツと対話するスクリプトを記述するのに使用できるオブジェクトが含まれています。Office 2013 およびそれ以降を対象、追加の場合は、これらの共通 Api を使用してください。 このオブジェクト モデルでは、コールバックを使用します。

- **ホスト固有の Api** - **2016 の Office**で導入された Api です。 このオブジェクト モデルでは、Office クライアントを使用して、Office の JavaScript Api の将来を表すときに表示される一般的なオブジェクトに対応するホストの特定の厳密に型指定されたオブジェクトを提供します。 現在、ホスト固有の Api には、Word の JavaScript API および Excel の JavaScript API が含まれます。

## <a name="supported-host-applications"></a>サポートされるホスト アプリケーション

- [Excel](overview/excel-add-ins-reference-overview.md)
- [OneNote](overview/onenote-add-ins-javascript-reference.md)
- [Outlook](requirement-sets/outlook-api-requirement-sets.md)
- [Visio](overview/visio-javascript-reference-overview.md)
- [Word](overview/word-add-ins-reference-overview.md)
- [共有 API](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> [PowerPoint およびプロジェクト](requirement-sets/powerpoint-and-project-note.md)は、JavaScript API を使用したアドインをサポートします。 ただし、現在持たないホストに固有の Api です。 共有 API を通じて、これらのホストと対話します。

[サポートされるホストとその他の要件](https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins)の詳細について説明します。

## <a name="open-api-specifications"></a>Open API の仕様

新しい Office アドイン用の API の設計と開発にあたり、[Open API の仕様](openspec.md) ページでこれらに対するフィードバックの提供が可能になります。パイプラインの新機能をご確認いただき、設計の仕様に関する情報をお寄せください。

## <a name="see-also"></a>関連項目

- [Office の JavaScript API のリファレンス](https://docs.microsoft.com/javascript/api/overview/office?view=office-js)