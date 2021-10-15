| クラス | フィールド | 説明 |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#layoutId)|新しいスライドに使用するスライド レイアウトの ID を指定します。|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#slideMasterId)|新しいスライドに使用するスライド マスターの ID を指定します。|
|[BulletFormat](/javascript/api/powerpoint/powerpoint.bulletformat)|[visible](/javascript/api/powerpoint/powerpoint.bulletformat#visible)|段落内の箇条書きが表示される場合に指定します。|
|[ParagraphFormat](/javascript/api/powerpoint/powerpoint.paragraphformat)|[bulletFormat](/javascript/api/powerpoint/powerpoint.paragraphformat#bulletFormat)|段落の箇条書き形式を表します。|
||[horizontalAlignment](/javascript/api/powerpoint/powerpoint.paragraphformat#horizontalAlignment)|段落の水平方向の配置を表します。|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#slideMasters)|プレゼンテーション内のオブジェクト `SlideMaster` のコレクションを返します。|
||[tags](/javascript/api/powerpoint/powerpoint.presentation#tags)|プレゼンテーションに添付されているタグのコレクションを返します。|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[delete()](/javascript/api/powerpoint/powerpoint.shape#delete__)|図形コレクションから図形を削除します。|
||[fill](/javascript/api/powerpoint/powerpoint.shape#fill)|この図形の塗りつぶしの書式設定を返します。|
||[height](/javascript/api/powerpoint/powerpoint.shape#height)|図形の高さをポイントで指定します。|
||[id](/javascript/api/powerpoint/powerpoint.shape#id)|図形の一意の ID を取得します。|
||[left](/javascript/api/powerpoint/powerpoint.shape#left)|図形の左側からスライドの左側までの距離をポイントで指定します。|
||[lineFormat](/javascript/api/powerpoint/powerpoint.shape#lineFormat)|この図形の線の書式設定を返します。|
||[name](/javascript/api/powerpoint/powerpoint.shape#name)|この図形の名前を指定します。|
||[tags](/javascript/api/powerpoint/powerpoint.shape#tags)|図形内のタグのコレクションを返します。|
||[textFrame](/javascript/api/powerpoint/powerpoint.shape#textFrame)|この図形のテキスト フレーム オブジェクトを返します。|
||[top](/javascript/api/powerpoint/powerpoint.shape#top)|図形の上端からスライドの上端までの距離をポイントで指定します。|
||[type](/javascript/api/powerpoint/powerpoint.shape#type)|この図形の種類を返します。|
||[width](/javascript/api/powerpoint/powerpoint.shape#width)|図形の幅をポイント単位で指定します。|
|[ShapeAddOptions](/javascript/api/powerpoint/powerpoint.shapeaddoptions)|[height](/javascript/api/powerpoint/powerpoint.shapeaddoptions#height)|図形の高さをポイントで指定します。|
||[left](/javascript/api/powerpoint/powerpoint.shapeaddoptions#left)|図形の左側からスライドの左側までの距離をポイントで指定します。|
||[top](/javascript/api/powerpoint/powerpoint.shapeaddoptions#top)|図形の上端からスライドの上端までの距離をポイントで指定します。|
||[width](/javascript/api/powerpoint/powerpoint.shapeaddoptions#width)|図形の幅をポイント単位で指定します。|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[addGeometricShape(geometricShapeType: PowerPoint.GeometricShapeType、オプション?: PowerPoint。ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addGeometricShape_geometricShapeType__options_)|スライドに幾何学的な図形を追加します。|
||[addLine(connectorType?: PowerPoint.ConnectorType、オプション?: PowerPoint。ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addLine_connectorType__options_)|スライドに線を追加します。|
||[addTextBox(text: string, options?: PowerPoint.ShapeAddOptions)](/javascript/api/powerpoint/powerpoint.shapecollection#addTextBox_text__options_)|指定したテキストをコンテンツとしてスライドにテキスト ボックスを追加します。|
||[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#getCount__)|コレクション内の図形の数を取得します。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getItem_key_)|一意の ID を使用して図形を取得します。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#getItemAt_index_)|コレクション内の 0 から始るインデックスを使用して図形を取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getItemOrNullObject_id_)|一意の ID を使用して図形を取得します。|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[ShapeFill](/javascript/api/powerpoint/powerpoint.shapefill)|[clear()](/javascript/api/powerpoint/powerpoint.shapefill#clear__)|この図形の塗りつぶしの書式設定をクリアします。|
||[foregroundColor](/javascript/api/powerpoint/powerpoint.shapefill#foregroundColor)|図形塗りつぶしの前景色を HTML カラー形式で表し、#RRGGBB 形式 ("FFA500"など) または名前付き HTML 色 ("オレンジ色" など) として表します。|
||[setSolidColor(color: string)](/javascript/api/powerpoint/powerpoint.shapefill#setSolidColor_color_)|図形の塗りつぶしの書式設定を均一な色に設定します。|
||[transparency](/javascript/api/powerpoint/powerpoint.shapefill#transparency)|塗りつぶしの透明度の割合を 0.0 (不透明) から 1.0 (クリア) の値として指定します。|
||[type](/javascript/api/powerpoint/powerpoint.shapefill#type)|図形の塗りつぶしの種類を返します。|
|[ShapeFont](/javascript/api/powerpoint/powerpoint.shapefont)|[bold](/javascript/api/powerpoint/powerpoint.shapefont#bold)|フォントの太字の状態を表します。|
||[color](/javascript/api/powerpoint/powerpoint.shapefont#color)|テキストの色の HTML カラー コード表現 (例: "#FF0000" は赤を表します)。|
||[italic](/javascript/api/powerpoint/powerpoint.shapefont#italic)|フォントの斜体の状態を表します。|
||[name](/javascript/api/powerpoint/powerpoint.shapefont#name)|フォント名 ("Calibri" など) を表します。|
||[size](/javascript/api/powerpoint/powerpoint.shapefont#size)|フォント サイズをポイント (11 など) で表します。|
||[underline](/javascript/api/powerpoint/powerpoint.shapefont#underline)|フォントに適用する下線の種類。|
|[ShapeLineFormat](/javascript/api/powerpoint/powerpoint.shapelineformat)|[color](/javascript/api/powerpoint/powerpoint.shapelineformat#color)|線の色を HTML カラー形式で表し、#RRGGBB 形式 ("FFA500" など) または名前付き HTML 色 ("オレンジ色" など) として表します。|
||[dashStyle](/javascript/api/powerpoint/powerpoint.shapelineformat#dashStyle)|線のダッシュ スタイルを表します。|
||[style](/javascript/api/powerpoint/powerpoint.shapelineformat#style)|図形の線スタイルを表します。|
||[transparency](/javascript/api/powerpoint/powerpoint.shapelineformat#transparency)|行の透明度のパーセンテージを 0.0 (不透明) から 1.0 (クリア) の値として指定します。|
||[visible](/javascript/api/powerpoint/powerpoint.shapelineformat#visible)|図形要素の線の書式設定が表示される場合に指定します。|
||[weight](/javascript/api/powerpoint/powerpoint.shapelineformat#weight)|線の太さ (ポイント数) を表します。|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[レイアウト](/javascript/api/powerpoint/powerpoint.slide#layout)|スライドのレイアウトを取得します。|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#shapes)|スライド内の図形のコレクションを返します。|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slideMaster)|スライドの `SlideMaster` 既定のコンテンツを表すオブジェクトを取得します。|
||[tags](/javascript/api/powerpoint/powerpoint.slide#tags)|スライド内のタグのコレクションを返します。|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint.AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_)|コレクションの最後に新しいスライドを追加します。|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#id)|スライド レイアウトの一意の ID を取得します。|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#name)|スライド レイアウトの名前を取得します。|
||[shapes](/javascript/api/powerpoint/powerpoint.slidelayout#shapes)|スライド レイアウト内の図形のコレクションを返します。|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getCount__)|コレクション内のレイアウトの数を取得します。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItem_key_)|一意の ID を使用してレイアウトを取得します。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItemAt_index_)|コレクション内の 0 から始るインデックスを使用してレイアウトを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItemOrNullObject_id_)|一意の ID を使用してレイアウトを取得します。|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#id)|スライド マスターの一意の ID を取得します。|
||[レイアウト](/javascript/api/powerpoint/powerpoint.slidemaster#layouts)|スライドのスライド マスターによって提供されるレイアウトのコレクションを取得します。|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#name)|スライド マスターの一意の名前を取得します。|
||[shapes](/javascript/api/powerpoint/powerpoint.slidemaster#shapes)|スライド マスターの図形のコレクションを返します。|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#getCount__)|コレクション内のスライド マスターの数を取得します。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItem_key_)|一意の ID を使用してスライド マスターを取得します。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItemAt_index_)|コレクション内の 0 から始るインデックスを使用してスライド マスターを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItemOrNullObject_id_)|一意の ID を使用してスライド マスターを取得します。|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#key)|タグの一意の ID を取得します。|
||[value](/javascript/api/powerpoint/powerpoint.tag#value)|タグの値を取得します。|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add(key: string, value: string)](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_)|コレクションの末尾に新しいタグを追加します。|
||[delete(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#delete_key_)|このコレクション内の指定されたタグ `key` を削除します。|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#getCount__)|コレクション内のタグの数を取得します。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getItem_key_)|一意の ID を使用してタグを取得します。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#getItemAt_index_)|コレクション内の 0 から始るインデックスを使用してタグを取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getItemOrNullObject_key_)|一意の ID を使用してタグを取得します。|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[TextFrame](/javascript/api/powerpoint/powerpoint.textframe)|[autoSizeSetting](/javascript/api/powerpoint/powerpoint.textframe#autoSizeSetting)|テキスト フレームの自動サイズ設定。|
||[bottomMargin](/javascript/api/powerpoint/powerpoint.textframe#bottomMargin)|テキスト フレームの下余白を表します (ポイント数)。|
||[deleteText()](/javascript/api/powerpoint/powerpoint.textframe#deleteText__)|テキスト フレーム内のテキストをすべて削除します。|
||[hasText](/javascript/api/powerpoint/powerpoint.textframe#hasText)|テキスト フレームにテキストが含まれている場合に指定します。|
||[leftMargin](/javascript/api/powerpoint/powerpoint.textframe#leftMargin)|テキスト フレームの左余白を表します (ポイント数)。|
||[rightMargin](/javascript/api/powerpoint/powerpoint.textframe#rightMargin)|テキスト フレームの右余白を表します (ポイント数)。|
||[textRange](/javascript/api/powerpoint/powerpoint.textframe#textRange)|テキスト フレーム内の図形にアタッチされているテキスト、およびテキストを操作するためのプロパティとメソッドを表します。|
||[topMargin](/javascript/api/powerpoint/powerpoint.textframe#topMargin)|テキスト フレームの上余白を表します (ポイント数)。|
||[verticalAlignment](/javascript/api/powerpoint/powerpoint.textframe#verticalAlignment)|テキスト フレームの垂直方向の配置を表します。|
||[wordWrap](/javascript/api/powerpoint/powerpoint.textframe#wordWrap)|図形内のテキストに合わせて線が自動的に折れるかどうかを指定します。|
|[TextRange](/javascript/api/powerpoint/powerpoint.textrange)|[font](/javascript/api/powerpoint/powerpoint.textrange#font)|テキスト範囲の `ShapeFont` フォント属性を表すオブジェクトを返します。|
||[getSubstring(start: number, length?: number)](/javascript/api/powerpoint/powerpoint.textrange#getSubstring_start__length_)|指定した範囲 `TextRange` の部分文字列のオブジェクトを返します。|
||[paragraphFormat](/javascript/api/powerpoint/powerpoint.textrange#paragraphFormat)|テキスト範囲の段落形式を表します。|
||[text](/javascript/api/powerpoint/powerpoint.textrange#text)|テキスト範囲のプレーン テキスト コンテンツを表します。|
