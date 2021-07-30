| クラス | フィールド | 説明 |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#onDataChanged)|コンテンツ コントロール内のデータが変更された場合に発生します。|
||[onDeleted](/javascript/api/word/word.contentcontrol#onDeleted)|コンテンツ コントロールが削除された場合に発生します。|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onSelectionChanged)|コンテンツ コントロール内の選択が変更された場合に発生します。|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentControl)|イベントを発生させたオブジェクト。|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventType)|イベントの種類。|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete__)|カスタム XML パーツを削除します。|
||[deleteAttribute(xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#deleteAttribute_xpath__namespaceMappings__name_)|xpath で識別される要素から、指定された名前の属性を削除します。|
||[deleteElement(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#deleteElement_xpath__namespaceMappings_)|xpath で識別される要素を削除します。|
||[getXml()](/javascript/api/word/word.customxmlpart#getXml__)|カスタム XML パーツの完全な XML コンテンツを取得します。|
||[insertAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#insertAttribute_xpath__namespaceMappings__name__value_)|指定された名前と値を持つ属性を、xpath で識別される要素に挿入します。|
||[insertElement(xpath: string, xml: string, namespaceMappings: any, index?: number)](/javascript/api/word/word.customxmlpart#insertElement_xpath__xml__namespaceMappings__index_)|xpath で識別される親要素の下に、指定された XML を子位置インデックスに挿入します。|
||[query(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#query_xpath__namespaceMappings_)|カスタム XML パーツの XML コンテンツを照会します。|
||[id](/javascript/api/word/word.customxmlpart#id)|カスタム XML パーツの ID を取得します。|
||[namespaceUri](/javascript/api/word/word.customxmlpart#namespaceUri)|カスタム XML パーツの名前空間 URI を取得します。|
||[setXml(xml: string)](/javascript/api/word/word.customxmlpart#setXml_xml_)|カスタム XML パーツの完全な XML コンテンツを設定します。|
||[updateAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#updateAttribute_xpath__namespaceMappings__name__value_)|xpath で識別される要素の指定された名前を持つ属性の値を更新します。|
||[updateElement(xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#updateElement_xpath__xml__namespaceMappings_)|xpath で識別される要素の XML を更新します。|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add(xml: string)](/javascript/api/word/word.customxmlpartcollection#add_xml_)|新しいカスタム XML パーツをドキュメントに追加します。|
||[getByNamespace(namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#getByNamespace_namespaceUri_)|名前空間が指定した名前空間に一致する、カスタム XML パーツの新しい範囲のコレクションを取得します。|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#getCount__)|コレクション内のアイテムの数を取得します。|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#getItem_id_)|ID に基づいて、カスタム XML パーツを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#getItemOrNullObject_id_)|ID に基づいて、カスタム XML パーツを取得します。|
||[items](/javascript/api/word/word.customxmlpartcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#getCount__)|コレクション内のアイテムの数を取得します。|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getItem_id_)|ID に基づいて、カスタム XML パーツを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getItemOrNullObject_id_)|ID に基づいて、カスタム XML パーツを取得します。|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#getOnlyItem__)|コレクションに含まれる項目が 1 つだけの場合、このメソッドはその項目を返します。|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#getOnlyItemOrNullObject__)|コレクションに含まれる項目が 1 つだけの場合、このメソッドはその項目を返します。|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Document](/javascript/api/word/word.document)|[deleteBookmark(name: string)](/javascript/api/word/word.document#deleteBookmark_name_)|ブックマークが存在する場合は、ドキュメントから削除します。|
||[getBookmarkRange(name: string)](/javascript/api/word/word.document#getBookmarkRange_name_)|ブックマークの範囲を取得します。|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.document#getBookmarkRangeOrNullObject_name_)|ブックマークの範囲を取得します。|
||[customXmlParts](/javascript/api/word/word.document#customXmlParts)|ドキュメント内のカスタム XML パーツを取得します。|
||[onContentControlAdded](/javascript/api/word/word.document#onContentControlAdded)|コンテンツ コントロールが追加された場合に発生します。|
||[settings](/javascript/api/word/word.document#settings)|ドキュメント内のアドインの設定を取得します。|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[deleteBookmark(name: string)](/javascript/api/word/word.documentcreated#deleteBookmark_name_)|ブックマークが存在する場合は、ドキュメントから削除します。|
||[getBookmarkRange(name: string)](/javascript/api/word/word.documentcreated#getBookmarkRange_name_)|ブックマークの範囲を取得します。|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.documentcreated#getBookmarkRangeOrNullObject_name_)|ブックマークの範囲を取得します。|
||[customXmlParts](/javascript/api/word/word.documentcreated#customXmlParts)|ドキュメント内のカスタム XML パーツを取得します。|
||[settings](/javascript/api/word/word.documentcreated#settings)|ドキュメント内のアドインの設定を取得します。|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#imageFormat)|インライン イメージの形式を取得します。|
|[List](/javascript/api/word/word.list)|[getLevelFont(level: number)](/javascript/api/word/word.list#getLevelFont_level_)|リスト内の指定されたレベルの箇条書き、数字、または図のフォントを取得します。|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#getLevelPicture_level_)|リスト内の指定されたレベルの図の base64 エンコードされた文字列表現を取得します。|
||[resetLevelFont(level: number, resetFontName?: boolean)](/javascript/api/word/word.list#resetLevelFont_level__resetFontName_)|箇条書き、番号、または図のフォントを、リスト内の指定されたレベルでリセットします。|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#setLevelPicture_level__base64EncodedImage_)|リスト内の指定されたレベルで図を設定します。|
|[Range](/javascript/api/word/word.range)|[getBookmarks(includeHidden?: boolean, includeAdjacent?: boolean)](/javascript/api/word/word.range#getBookmarks_includeHidden__includeAdjacent_)|範囲内または範囲に重なるすべてのブックマークの名前を取得します。|
||[insertBookmark(name: string)](/javascript/api/word/word.range#insertBookmark_name_)|範囲にブックマークを挿入します。|
|[設定](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete__)|設定を削除します。|
||[key](/javascript/api/word/word.setting#key)|設定のキーを取得します。|
||[value](/javascript/api/word/word.setting#value)|設定の値を取得または設定します。|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add(key: string, value: any)](/javascript/api/word/word.settingcollection#add_key__value_)|新しい設定を作成するか、既存の設定を設定します。|
||[deleteAll()](/javascript/api/word/word.settingcollection#deleteAll__)|このアドインのすべての設定を削除します。|
||[getCount()](/javascript/api/word/word.settingcollection#getCount__)|設定の数を取得します。|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getItem_key_)|キーによって設定オブジェクトを取得します。大文字と小文字が区別されます。|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getItemOrNullObject_key_)|キーによって設定オブジェクトを取得します。大文字と小文字が区別されます。|
||[items](/javascript/api/word/word.settingcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[表](/javascript/api/word/word.table)|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/javascript/api/word/word.table#mergeCells_topRow__firstCell__bottomRow__lastCell_)|最初のセルと最後のセルで結合されたセルを結合します。|
|[TableCell](/javascript/api/word/word.tablecell)|[split(rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#split_rowCount__columnCount_)|セルを指定した数の行と列に分割します。|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertContentControl__)|行にコンテンツ コントロールを挿入します。|
||[merge()](/javascript/api/word/word.tablerow#merge__)|行を 1 つのセルに結合します。|
