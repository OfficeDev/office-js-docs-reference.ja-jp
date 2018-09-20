# <a name="defaultsettings-element"></a>DefaultSettings 要素

既定のソースの場所と、コンテンツの他の既定の設定を指定または作業ウィンドウ アドインです。

**アドインの種類:** コンテンツ、作業ウィンドウ

## <a name="syntax"></a>構文

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a>含まれています。

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>含めることができます。

|**要素**|**コンテンツ**|**メール**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](sourcelocation.md)|x||x|
|[RequestedWidth](requestedwidth.md)|x|||
|[RequestedHeight](requestedheight.md)|x|||

## <a name="remarks"></a>注釈

**DefaultSettings** 要素のソースの場所と他の設定が適用されるのは、コンテンツ アドインと作業ウィンドウ アドインのみです。メール アドインの場合は、ソース ファイルの既定の場所とその他の既定の設定を [FormSettings](formsettings.md) 要素に指定します。

