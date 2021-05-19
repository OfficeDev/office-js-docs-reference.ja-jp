| クラス | フィールド | 説明 |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#layoutid)|新しいスライドに使用するスライド レイアウトの ID を指定します。|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#slidemasterid)|新しいスライドに使用するスライド マスターの ID を指定します。|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#slidemasters)|プレゼンテーション内のオブジェクト `SlideMaster` のコレクションを返します。|
||[tags](/javascript/api/powerpoint/powerpoint.presentation#tags)|プレゼンテーションに添付されているタグのコレクションを返します。|
|[図形](/javascript/api/powerpoint/powerpoint.shape)|[delete()](/javascript/api/powerpoint/powerpoint.shape#delete--)|図形コレクションから図形を削除します。|
||[id](/javascript/api/powerpoint/powerpoint.shape#id)|図形の一意の ID を取得します。|
||[tags](/javascript/api/powerpoint/powerpoint.shape#tags)|図形内のタグのコレクションを返します。|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#getcount--)|コレクション内の図形の数を取得します。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getitem-key-)|一意の ID を使用して図形を取得します。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemat-index-)|コレクション内の 0 から始るインデックスを使用して図形を取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getitemornullobject-id-)|一意の ID を使用して図形を取得します。|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[レイアウト](/javascript/api/powerpoint/powerpoint.slide#layout)|スライドのレイアウトを取得します。|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#shapes)|スライド内の図形のコレクションを返します。|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slidemaster)|スライドの `SlideMaster` 既定のコンテンツを表すオブジェクトを取得します。|
||[tags](/javascript/api/powerpoint/powerpoint.slide#tags)|スライド内のタグのコレクションを返します。|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint.AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#add-options-)|コレクションの最後に新しいスライドを追加します。|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#id)|スライド レイアウトの一意の ID を取得します。|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#name)|スライド レイアウトの名前を取得します。|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getcount--)|コレクション内のレイアウトの数を取得します。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitem-key-)|一意の ID を使用してレイアウトを取得します。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemat-index-)|コレクション内の 0 から始るインデックスを使用してレイアウトを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getitemornullobject-id-)|一意の ID を使用してレイアウトを取得します。|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#id)|スライド マスターの一意の ID を取得します。|
||[レイアウト](/javascript/api/powerpoint/powerpoint.slidemaster#layouts)|スライドのスライド マスターによって提供されるレイアウトのコレクションを取得します。|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#name)|スライド マスターの一意の名前を取得します。|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#getcount--)|コレクション内のスライド マスターの数を取得します。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitem-key-)|一意の ID を使用してスライド マスターを取得します。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemat-index-)|コレクション内の 0 から始るインデックスを使用してスライド マスターを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getitemornullobject-id-)|一意の ID を使用してスライド マスターを取得します。|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#key)|タグの一意の ID を取得します。|
||[value](/javascript/api/powerpoint/powerpoint.tag#value)|タグの値を取得します。|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add(key: string, value: string)](/javascript/api/powerpoint/powerpoint.tagcollection#add-key--value-)|コレクションの末尾に新しいタグを追加します。|
||[delete(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#delete-key-)|このコレクション内の指定されたタグ `key` を削除します。|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#getcount--)|コレクション内のタグの数を取得します。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getitem-key-)|一意の ID を使用してタグを取得します。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#getitemat-index-)|コレクション内の 0 から始るインデックスを使用してタグを取得します。|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getitemornullobject-key-)|一意の ID を使用してタグを取得します。|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#items)|このコレクション内に読み込まれた子アイテムを取得します。|
