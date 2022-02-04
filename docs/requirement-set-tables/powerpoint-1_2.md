| クラス | フィールド | 説明 |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[書式設定](/javascript/api/powerpoint/powerpoint.insertslideoptions#powerpoint-powerpoint-insertslideoptions-formatting-member)|スライド挿入時に使用する書式を指定します。|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#powerpoint-powerpoint-insertslideoptions-sourceslideids-member)|現在のプレゼンテーションに挿入されるソース プレゼンテーションのスライドを指定します。|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#powerpoint-powerpoint-insertslideoptions-targetslideid-member)|プレゼンテーション内で新しいスライドを挿入する場所を指定します。|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64(base64File: string, options?: PowerPoint.InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-insertslidesfrombase64-member(1))|指定したスライドをプレゼンテーションから現在のプレゼンテーションに挿入します。|
||[スライド](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-slides-member)|プレゼンテーション内のスライドの順序付きコレクションを返します。|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-delete-member(1))|プレゼンテーションからスライドを削除します。|
||[id](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-id-member)|スライドの一意の ID を取得します。|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getcount-member(1))|コレクション内のスライドの数を取得します。|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitem-member(1))|一意の ID を使用してスライドを取得します。|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitemat-member(1))|コレクション内の 0 から始るインデックスを使用してスライドを取得します。|
||[getItemOrNullObject(id: string)](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitemornullobject-member(1))|一意の ID を使用してスライドを取得します。|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-items-member)|このコレクション内に読み込まれた子アイテムを取得します。|
