# <a name="iconurl-element"></a>IconUrl 要素

挿入 UX と Office ストアの Office アドインを表すために使用されるイメージの URL を指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a>含めることができます。

[Override](override.md)

## <a name="attributes"></a>属性

|**属性**|**型**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|DefaultValue|文字列|必須|この設定の既定値を指定します。この値は、[DefaultLocale](defaultlocale.md) 要素に指定されるロケールを対象としています。|

## <a name="remarks"></a>備考

メール アドインの場合、アイコンは、**[ファイル]**  >  **[アドインの管理]** UI (Outlook) または **[設定]**  >  **[アドインの管理]** UI (Outlook Web App) に表示されます。コンテンツ アドインまたは作業ウィンドウ アドインでは、アイコンは、**[挿入]**  >  **[アドイン]** UI に表示されます。どのアドインの種類についても、アドインを Office ストアに公開すると、アイコンは Office ストア サイトでも使用されます。

イメージは次のファイル形式のいずれかである必要があります: GIF、JPG、PNG、EXIF、BMP や TIFF です。 コンテンツと作業ウィンドウ アプリでは、指定したイメージは 32 x 32 ピクセルである必要があります。 メール アプリケーションでは、イメージは 64 × 64 ピクセルである必要があります。 [HighResolutionIconUrl](highresolutioniconurl.md)要素を使用して高 DPI の画面で実行して、Office ホスト アプリケーションで使用するアイコンを指定することもする必要があります。 詳細については、 [AppSource で、Office 内で効果的な一覧を作成する](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity)に _、アプリケーションの一貫性のあるビジュアルを作成_する」を参照してください。
