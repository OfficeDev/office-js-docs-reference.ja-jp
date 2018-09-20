# <a name="set-element"></a>Set 要素

Office アドインをアクティブにするために必要な JavaScript API for Office の要件セットを指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>含まれています。

[Sets](sets.md)

## <a name="attributes"></a>属性

|**属性**|**型**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|名前|文字列|必須|[要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)の名前。|
|MinVersion|文字列|省略可能|アドインに必要な API セットの最小バージョンを指定します。**DefaultMinVersion** の値が親の [Sets](sets.md) 要素に指定されている場合は、その値を上書きします。|

## <a name="remarks"></a>備考

要求セットの詳細については、 [Office のバージョンおよび要件の設定](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)を参照してください。

**Set** 要素の **MinVersion** 属性と **Sets** 要素の **DefaultMinVersion** 属性の詳細については、「[マニフェストで Requirements 要素を設定する](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)」をご覧ください。

> [!IMPORTANT] 
> メール アドインの場合、1 つだけです`"Mailbox"`利用可能な要件を設定します。 この要件のセットには、outlook でメールのアドインでサポートされている API の全体のサブセットが含まれているし、指定する必要があります、`"Mailbox"`要件は、メールでこのアドインのマニフェストの設定 (は省略可能な場合と同様の内容とタスクのウィンドウ - アドイン)。 また、メールのアドインの特定のメソッドのサポートを宣言できません。
