# <a name="method-element"></a>Method 要素

Office アドインをアクティブにするために必要な JavaScript API for Office の個別のメソッドを指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ

## <a name="syntax"></a>構文

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>含まれています。

[メソッド](methods.md)

## <a name="attributes"></a>属性

|**属性**|**型**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|名前|文字列|必須|必要なメソッドの名前をその親オブジェクトで修飾して指定します。たとえば、**getSelectedDataAsync** メソッドを指定するには、`"Document.getSelectedDataAsync"` と指定する必要があります。|

## <a name="remarks"></a>備考

メールのアドインでは、**メソッド**および**メソッド**の要素はサポートされていません。要求セットの詳細については、 [Office のバージョンおよび要件の設定](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)を参照してください。

> [!IMPORTANT] 
> 個々 のメソッドの最小バージョン要件を指定する方法がないため、メソッドが、実行時に使用可能であることを確認する必要がありますもを使用する**if**ステートメントの追加のスクリプトでそのメソッドを呼び出すときにします。 これを行う方法の詳細については、 [Office 用の JavaScript API を理解する](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)を参照してください。

