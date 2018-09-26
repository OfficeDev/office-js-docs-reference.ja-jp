
# <a name="item"></a>item

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|制限あり|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成または読み取り|

##### <a name="members-and-methods"></a>メンバーとメソッド

| メンバー | 種類 |
|--------|------|
| [attachments](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | メンバー |
| [bcc](#bcc-recipientsjavascriptapioutlookofficerecipients) | メンバー |
| [body](#body-bodyjavascriptapioutlookofficebody) | メンバー |
| [cc](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | メンバー |
| [conversationId](#nullable-conversationid-string) | メンバー |
| [dateTimeCreated](#datetimecreated-date) | メンバー |
| [dateTimeModified](#datetimemodified-date) | メンバー |
| [end](#end-datetimejavascriptapioutlookofficetime) | メンバー |
| [from](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | メンバー |
| [internetMessageId](#internetmessageid-string) | メンバー |
| [itemClass](#itemclass-string) | メンバー |
| [itemId](#nullable-itemid-string) | メンバー |
| [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | メンバー |
| [location](#location-stringlocationjavascriptapioutlookofficelocation) | メンバー |
| [normalizedSubject](#normalizedsubject-string) | メンバー |
| [notificationMessages](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | メンバー |
| [optionalAttendees](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | メンバー |
| [organizer](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | メンバー |
| [recurrence](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | メンバー |
| [requiredAttendees](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | メンバー |
| [sender](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | メンバー |
| [seriesId](#nullable-seriesid-string) | メンバー |
| [start](#start-datetimejavascriptapioutlookofficetime) | メンバー |
| [subject](#subject-stringsubjectjavascriptapioutlookofficesubject) | メンバー |
| [to](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | メンバー |
| [addFileAttachmentAsync](#addfileattachmentasyncuri-attachmentname-options-callback) | メソッド |
| [addFileAttachmentFromBase64Async](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | メソッド |
| [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) | メソッド |
| [addItemAttachmentAsync](#additemattachmentasyncitemid-attachmentname-options-callback) | メソッド |
| [close](#close) | メソッド |
| [displayReplyAllForm](#displayreplyallformformdata) | メソッド |
| [displayReplyForm](#displayreplyformformdata) | メソッド |
| [getEntities](#getentities--entitiesjavascriptapioutlookofficeentities) | メソッド |
| [getEntitiesByType](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | メソッド |
| [getFilteredEntitiesByName](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | メソッド |
| [getInitializationContextAsync](#getinitializationcontextasyncoptions-callback) | メソッド |
| [getRegExMatches](#getregexmatches--object) | メソッド |
| [getRegExMatchesByName](#getregexmatchesbynamename--nullable-array-string-) | メソッド |
| [getSelectedDataAsync](#getselecteddataasynccoerciontype-options-callback--string) | メソッド |
| [getSelectedEntities](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | メソッド |
| [getSelectedRegExMatches](#getselectedregexmatches--object) | メソッド |
| [getSharedPropertiesAsync](#getsharedpropertiesasyncoptions-callback) | メソッド |
| [loadCustomPropertiesAsync](#loadcustompropertiesasynccallback-usercontext) | メソッド |
| [removeAttachmentAsync](#removeattachmentasyncattachmentid-options-callback) | メソッド |
| [removeHandlerAsync](#removehandlerasynceventtype-handler-options-callback) | メソッド |
| [saveAsync](#saveasyncoptions-callback) | メソッド |
| [setSelectedDataAsync](#setselecteddataasyncdata-options-callback) | メソッド |

### <a name="example"></a>例

次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。

```
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a>メンバー

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a>attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)>

アイテムの添付ファイルの配列を取得します。閲覧モードのみ。

> [!NOTE]
> ファイルの特定の種類は、潜在的なセキュリティの問題により、Outlook によってブロックされは返されません。 詳細については、 [Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)を参照してください。

##### <a name="type"></a>型:

*   Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)>

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|読み取り|

##### <a name="example"></a>例

次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。

```
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a>bcc :[Recipients](/javascript/api/outlook/office.recipients)

取得またはメッセージの bcc (ブラインド カーボン コピー) 受信者を更新するメソッドを提供するオブジェクトを取得します。 新規作成モードのみ。

##### <a name="type"></a>型:

*   [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.1|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成|

##### <a name="example"></a>例

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a>body :[Body](/javascript/api/outlook/office.body)

アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。

##### <a name="type"></a>型:

*   [Body](/javascript/api/outlook/office.body)

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.1|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成または読み取り|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>[cc]: 配列 <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[の受信者](/javascript/api/outlook/office.recipients)。

メッセージの Cc (カーボン コピー) 受信者へのアクセスを提供します。 オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。

##### <a name="read-mode"></a>閲覧モード

`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。

##### <a name="compose-mode"></a>新規作成モード

`cc`を`Recipients`オブジェクトを取得または、メッセージの**Cc**行の受信者を更新するメソッドを提供します。

##### <a name="type"></a>型:

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成または読み取り|

##### <a name="example"></a>例

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a>(nullable) conversationId :String

特定のメッセージが含まれている電子メールの会話の識別子を取得します。

メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。

新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。

##### <a name="type"></a>型:

*   String

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成または読み取り|

#### <a name="datetimecreated-date"></a>dateTimeCreated :Date

アイテムが作成された日時を取得します。閲覧モードのみ。

##### <a name="type"></a>型:

*   日付

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|読み取り|

##### <a name="example"></a>例

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a>dateTimeModified :Date

アイテムが最後に変更された日時を取得します。閲覧モードのみ。

> [!NOTE]
> IOS は、Outlook または Outlook Android のでは、このメンバーはサポートされていません。

##### <a name="type"></a>型:

*   日付

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|読み取り|

##### <a name="example"></a>例

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a>end :Date|[Time](/javascript/api/outlook/office.time)

予定が終了する日時を取得または設定します。

`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。

##### <a name="read-mode"></a>閲覧モード

`end` プロパティは `Date` オブジェクトを返します。

##### <a name="compose-mode"></a>新規作成モード

`end` プロパティは `Time` オブジェクトを返します。

[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。

##### <a name="type"></a>型:

*   Date | [Time](/javascript/api/outlook/office.time)

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成または読み取り|

##### <a name="example"></a>例

次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。

```
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a>:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[から](/javascript/api/outlook/office.from)

メッセージの送信者の電子メール アドレスを取得します。

メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。

> [!NOTE]
> `recipientType`のプロパティの`EmailAddressDetails`オブジェクトで、`from`プロパティは、 `undefined`。

##### <a name="read-mode"></a>閲覧モード

`from`を`EmailAddressDetails`オブジェクトです。

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a>新規作成モード

`from`を`From`を取得するメソッドを提供するオブジェクト、値からです。

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a>型:

*   [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [から](/javascript/api/outlook/office.from)

##### <a name="requirements"></a>要件

|要件|||
|---|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|1.7|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|ReadWriteItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Read|作成|

#### <a name="internetmessageid-string"></a>internetMessageId :String

電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。

##### <a name="type"></a>型:

*   String

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|読み取り|

##### <a name="example"></a>例

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a>itemClass :String

選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。

`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。

|種類|説明|アイテム クラス|
|---|---|---|
|予定アイテム|アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurence` の予定表アイテムは次のとおりです。|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|メッセージ アイテム|これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など) を作成できます。

##### <a name="type"></a>型:

*   String

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|読み取り|

##### <a name="example"></a>例

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a>(nullable) itemId :String

現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。

> [!NOTE]
> `itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。 `itemId`プロパティは、Outlook のエントリ ID または Outlook の REST API によって使用される ID と同じではありません。 この値を使用して REST API の呼び出しを行う前にすると[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)を使用してを変換する必要があります。 詳細については、 [Outlook のアドインから Outlook REST Api の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)を参照してください。

新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。

##### <a name="type"></a>種類:

*   String

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|読み取り|

##### <a name="example"></a>例

次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a>itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)

インスタンスが表しているアイテムの種類を取得します。

`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。

##### <a name="type"></a>型:

*   [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成または読み取り|

##### <a name="example"></a>例

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a>location :String|[Location](/javascript/api/outlook/office.location)

予定の場所を取得または設定します。

##### <a name="read-mode"></a>閲覧モード

`location` プロパティは、予定の場所を格納した文字列を返します。

##### <a name="compose-mode"></a>新規作成モード

`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。

##### <a name="type"></a>型:

*   String | [Location](/javascript/api/outlook/office.location)

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成または読み取り|

##### <a name="example"></a>例

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a>normalizedSubject :String

すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。

normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) プロパティを使用します。

##### <a name="type"></a>型:

*   String

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|読み取り|

##### <a name="example"></a>例

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a>notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)

アイテムの通知メッセージを取得します。

##### <a name="type"></a>型:

*   [NotificationMessages](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.3|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成または読み取り|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)

イベントの任意の出席者へのアクセスを提供します。 オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。

##### <a name="read-mode"></a>閲覧モード

`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。

##### <a name="compose-mode"></a>新規作成モード

`optionalAttendees`を`Recipients`オブジェクトを取得または省略可能な会議の出席者を更新するメソッドを提供します。

##### <a name="type"></a>型:

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成または読み取り|

##### <a name="example"></a>例

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a>オーガナイザー:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[オーガナイザー](/javascript/api/outlook/office.organizer)

指定した会議の開催者の電子メール アドレスを取得します。

##### <a name="read-mode"></a>閲覧モード

`organizer`プロパティは、会議の開催者を表す[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)オブジェクトを返します。

##### <a name="compose-mode"></a>新規作成モード

`organizer`プロパティが開催者の値を取得するメソッドを提供する[構成内容変更](/javascript/api/outlook/office.organizer)オブジェクトを返します。

##### <a name="type"></a>型:

*   [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [オーガナイザー](/javascript/api/outlook/office.organizer)

##### <a name="requirements"></a>要件

|要件|||
|---|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|1.7|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|ReadWriteItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|Read|作成|

##### <a name="example"></a>例

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a>(許容) 定期的:[定期的なアイテム](/javascript/api/outlook/office.recurrence)

取得または予定の定期的なパターンを設定します。 定期的な会議出席依頼を取得します。 モードの予定表アイテムを読んだり作成したりします。 会議出席依頼アイテムの読み取りモードです。

`recurrence`プロパティは、アイテムが系列または系列のインスタンスである場合に定期的な予定または会議出席依頼に[定期的なアイテム](/javascript/api/outlook/office.recurrence)オブジェクトを返します。 `null`単独の予定および会議出席依頼を単独の予定が返されます。 `undefined`会議出席依頼ではないメッセージが返されます。

> 注: 会議出席依頼がある、 `itemClass` IPM の値です。Schedule.Meeting.Request。

> 注: 定期的なアイテム オブジェクトがある場合`null`、これは、オブジェクトが 1 つの予定または会議出席依頼、単独の予定および一連の一部ではないのであることを示します。

##### <a name="type"></a>型:

* [定期的なアイテム](/javascript/api/outlook/office.recurrence)

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.7|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成または読み取り|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)

イベントの出席者へのアクセスを提供します。 オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。

##### <a name="read-mode"></a>閲覧モード

`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。

##### <a name="compose-mode"></a>新規作成モード

`requiredAttendees`を`Recipients`オブジェクトを取得または会議の出席者を更新するメソッドを提供します。

##### <a name="type"></a>型:

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成または読み取り|

##### <a name="example"></a>例

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a>sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)

電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。

メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。

> [!NOTE]
> `recipientType`のプロパティの`EmailAddressDetails`オブジェクトで、`sender`プロパティは、 `undefined`。

##### <a name="type"></a>型:

*   [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|読み取り|

##### <a name="example"></a>例

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a>(許容) seriesId: 文字列

インスタンスが属する系列の id を取得します。

OWA と outlook 2002 で、`seriesId`は、この項目が属する親 (系列) アイテムの Exchange Web サービス (EWS) の ID を返します。 IOS および Android で、 `seriesId` 、親項目の残りの部分 ID を返します。

> [!NOTE]
> `seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。 `seriesId`プロパティは Outlook の REST API で使用される Outlook の Id と同じではありません。 この値を使用して REST API の呼び出しを行う前にすると[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)を使用してを変換する必要があります。 詳細については、 [Outlook のアドインから Outlook REST Api の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api)を参照してください。

`seriesId`プロパティを返します。`null`アイテムの親アイテムを次のようにされていない単一の関連するアイテム、予定または会議を要求し、返しますの`undefined`、その他の項目の要求を満たしていません。

##### <a name="type"></a>型:

* String

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.7|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成または読み取り|

##### <a name="example"></a>例

```
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a>start :Date|[Time](/javascript/api/outlook/office.time)

予定を開始する日時を取得または設定します。

`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。

##### <a name="read-mode"></a>閲覧モード

`start` プロパティは `Date` オブジェクトを返します。

##### <a name="compose-mode"></a>新規作成モード

`start` プロパティは `Time` オブジェクトを返します。

[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。

##### <a name="type"></a>型:

*   Date | [Time](/javascript/api/outlook/office.time)

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成または読み取り|

##### <a name="example"></a>例

次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。

```
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a>subject :String|[Subject](/javascript/api/outlook/office.subject)

アイテムの件名フィールドに示される説明を取得または設定します。

`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。

##### <a name="read-mode"></a>閲覧モード

`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a>新規作成モード

`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a>型:

*   String | [Subject](/javascript/api/outlook/office.subject)

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成または読み取り|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>: 配列 <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[の受信者](/javascript/api/outlook/office.recipients)。

[メッセージの [**宛先**] 行の受信者へのアクセスを提供します。 オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。

##### <a name="read-mode"></a>閲覧モード

`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。

##### <a name="compose-mode"></a>新規作成モード

`to`を`Recipients`オブジェクトを取得または、メッセージの [**宛先**] 行の受信者を更新するメソッドを提供します。

##### <a name="type"></a>型:

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成または読み取り|

##### <a name="example"></a>例

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a>メソッド

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a>addFileAttachmentAsync(uri, attachmentName, [options], [callback])

ファイルを添付ファイルとしてメッセージまたは予定に追加します。

`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。

その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。

##### <a name="parameters"></a>パラメーター:
|名前|型|属性|説明|
|---|---|---|---|
|`uri`|String||メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。|
|`attachmentName`|String||添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。|
|`options`|Object|&lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`|Object|&lt;optional&gt;|開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。|
|`options.isInline`|Boolean|&lt;optional&gt;|`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。|
|`callback`|function|&lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。 <br/>成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。<br/>添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。|

##### <a name="errors"></a>エラー

|エラー コード|説明|
|------------|-------------|
|`AttachmentSizeExceeded`|添付ファイルのサイズが上限を超えています。|
|`FileTypeNotSupported`|許可されていない拡張子の添付ファイルです。|
|`NumberOfAttachmentsExceeded`|メッセージまたは予定の添付ファイルが多すぎます。|

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.1|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成|

##### <a name="examples"></a>例

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        
      }
    );
  }
);
```

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>addFileAttachmentFromBase64Async (base64File、attachmentName、[オプション]、[コールバック])

メッセージまたは予定を添付ファイルとしてエンコード base64 からファイルを追加します。

`addFileAttachmentFromBase64Async`メソッドは、base64 エンコーディングからファイルをアップロードし、作成フォーム内の項目にアタッチします。 このメソッドは、AsyncResult.value オブジェクトの添付ファイルの識別子を返します。

その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。

##### <a name="parameters"></a>パラメーター:
|名前|型|属性|説明|
|---|---|---|---|
|`base64File`|String||イメージや、電子メール、またはイベントに追加するファイルのコンテンツを base64 にエンコードされます。|
|`attachmentName`|String||添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。|
|`options`|Object|&lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`|Object|&lt;optional&gt;|開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。|
|`options.isInline`|Boolean|&lt;optional&gt;|`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。|
|`callback`|function|&lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。 <br/>成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。<br/>添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。|

##### <a name="errors"></a>エラー

|エラー コード|説明|
|------------|-------------|
|`AttachmentSizeExceeded`|添付ファイルのサイズが上限を超えています。|
|`FileTypeNotSupported`|許可されていない拡張子の添付ファイルです。|
|`NumberOfAttachmentsExceeded`|メッセージまたは予定の添付ファイルが多すぎます。|

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|プレビュー|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成|

##### <a name="examples"></a>例

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
      }
    );
  }
);
```

####  <a name="addhandlerasynceventtype-handler-options-callback"></a>addHandlerAsync(eventType, handler, [options], [callback])

サポートされているイベントのイベント ハンドラーを追加します。

現在サポートされているイベントの種類は、 `Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged`と`Office.EventType.RecurrenceChanged`

##### <a name="parameters"></a>パラメーター:

| 名前 | 型 | 属性 | 説明 |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || ハンドラーを呼び出す必要のあるイベント。 |
| `handler` | Function || イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。 |
| `options` | Object | &lt;optional&gt; | 次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。 |
| `options.asyncContext` | Object | &lt;optional&gt; | 開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。 |
| `callback` | function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.7 |
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a>addItemAttachmentAsync(itemId, attachmentName, [options], [callback])

メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。

`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。

その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。

Office アドインは、Outlook Web App で実行されている場合、`addItemAttachmentAsync`メソッドが項目を編集しているアイテム以外のアイテムに関連付けることができますただし、これはサポートされていません、お勧めできません。

##### <a name="parameters"></a>パラメーター:

|名前|型|属性|説明|
|---|---|---|---|
|`itemId`|String||添付するアイテムの Exchange 識別子。最大長は 100 文字です。|
|`attachmentName`|String||添付するアイテムの件名。最大長は 255 文字です。|
|`options`|Object|&lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`|Object|&lt;optional&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`|function|&lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。 <br/>成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。<br/>添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。|

##### <a name="errors"></a>エラー

|エラー コード|説明|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|メッセージまたは予定の添付ファイルが多すぎます。|

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.1|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成|

##### <a name="example"></a>例

次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。

```
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a>close()

作成中の現在の項目を閉じます。

`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。

> [!NOTE]
> アイテム予定は、以前保存されたを使用する場合は、web 上の Outlook で`saveAsync`を求めるメッセージを保存、破棄、または、キャンセル場合でも、変更が発生していないから、項目を保存します。

Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.3|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|制限あり|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成|

#### <a name="displayreplyallformformdata"></a>displayReplyAllForm(formData)

選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。

> [!NOTE]
> IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。

Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。

文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。

`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。

##### <a name="parameters"></a>パラメーター:

|名前|型|属性|説明|
|---|---|---|---|
|`formData`|String &#124; Object||回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。<br/>**または**<br/>本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。|
|`formData.htmlBody`|String|&lt;省略可能&gt;|回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。
|`formData.attachments`|Array.&lt;Object&gt;|&lt;省略可能&gt;|ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。|
|`formData.attachments.type`|String||添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。|
|`formData.attachments.name`|String||添付ファイル名を含む文字列。最大の長さは 255 文字です。|
|`formData.attachments.url`|String||`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。|
|`formData.attachments.isInline`|Boolean||`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。|
|`formData.attachments.itemId`|String||`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。|
|`callback`|function|&lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。|

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|読み取り|

##### <a name="examples"></a>例

次のコードは `displayReplyAllForm` 関数に文字列を渡します。

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

空の本文を返信します。

```
Office.context.mailbox.item.displayReplyAllForm({});
```

本文だけを返信します。

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

本文とファイルの添付ファイルを返信します。

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

本文とアイテムの添付ファイルを返信します。

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a>displayReplyForm(formData)

選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。

> [!NOTE]
> IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。

Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。

文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。

`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。

##### <a name="parameters"></a>パラメーター:

|名前|型|属性|説明|
|---|---|---|---|
|`formData`|String &#124; Object||回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。<br/>**または**<br/>本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。|
|`formData.htmlBody`|String|&lt;省略可能&gt;|回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。
|`formData.attachments`|Array.&lt;Object&gt;|&lt;省略可能&gt;|ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。|
|`formData.attachments.type`|String||添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。|
|`formData.attachments.name`|String||添付ファイル名を含む文字列。最大の長さは 255 文字です。|
|`formData.attachments.url`|String||`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。|
|`formData.attachments.isInline`|Boolean||`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。|
|`formData.attachments.itemId`|String||`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。|
|`callback`|function|&lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。|

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|読み取り|

##### <a name="examples"></a>例

次のコードは `displayReplyForm` 関数に文字列を渡します。

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

空の本文を返信します。

```
Office.context.mailbox.item.displayReplyForm({});
```

本文だけを返信します。

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

本文とファイルの添付ファイルを返信します。

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

本文とアイテムの添付ファイルを返信します。

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a>getEntities() → {[Entities](/javascript/api/outlook/office.entities)}

選択したアイテムの本文内のエンティティを取得します。

> [!NOTE]
> IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|読み取り|

##### <a name="returns"></a>戻り値:

型:[Entities](/javascript/api/outlook/office.entities)

##### <a name="example"></a>例

次の使用例は、現在の項目の本文に連絡先のエンティティを取得します。

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a>getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}

選択したアイテムの本文に指定されたエンティティ型のすべてのエンティティの配列を取得します。

> [!NOTE]
> IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。

##### <a name="parameters"></a>パラメーター:

|名前|種類|説明|
|---|---|---|
|`entityType`|[Office.MailboxEnums.EntityType](/javascript/api/outlook/office.mailboxenums.entitytype)|EntityType 列挙値の 1 つ。|

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|制限あり|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|読み取り|

##### <a name="returns"></a>戻り値:

`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。 アイテムの本文に指定した型のエンティティがない場合は、メソッドは空の配列を返します。 それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。

このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。

|`entityType` の値|返される配列内のオブジェクトの型|必要なアクセス許可のレベル|
|---|---|---|
|`Address`|文字列|**制限あり**|
|`Contact`|連絡先|**ReadItem**|
|`EmailAddress`|文字列|**ReadItem**|
|`MeetingSuggestion`|MeetingSuggestion|**ReadItem**|
|`PhoneNumber`|PhoneNumber|**制限あり**|
|`TaskSuggestion`|TaskSuggestion|**ReadItem**|
|`URL`|文字列|**制限あり**|

型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>

##### <a name="example"></a>例

次の例では、現在の項目の本文に郵便番号のアドレスを表す文字列の配列にアクセスする方法を示します。

```
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a>getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}

マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。

> [!NOTE]
> IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。

`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。

##### <a name="parameters"></a>パラメーター:

|名前|種類|説明|
|---|---|---|
|`name`|String|一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。|

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|読み取り|

##### <a name="returns"></a>戻り値:

`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。

型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>

#### <a name="getinitializationcontextasyncoptions-callback"></a>getInitializationContextAsync([options], [callback])

アドインが[操作可能メッセージによってアクティブ化](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを取得します。

> [!NOTE]
> このメソッドは Outlook 2016 または Windows (クイック実行バージョン 16.0.8413.1000 以降) と、web 上で Outlook を後で Office 365 のです。

##### <a name="parameters"></a>パラメーター:
|名前|型|属性|説明|
|---|---|---|---|
|`options`|オブジェクト|&lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`|Object|&lt;optional&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`|function|&lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。 <br/>成功した場合、初期化データが提供されている、`asyncResult.value`文字列としてのプロパティです。<br/>初期化コンテキストがない場合、`asyncResult` オブジェクトには、`code` プロパティが `9020`、`name` プロパティが `GenericResponseError` に設定された `Error` オブジェクトが含まれます。|

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|プレビュー|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|読み取り|

##### <a name="example"></a>例

```
// Get the initialization context (if present)
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object
        var context = JSON.parse(asyncResult.value);
        // Do something with context
      } else {
        // Empty context, treat as no context
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is
        // no context
        // Treat as no context
      } else {
        // Handle the error
      }
    }
  }
);
```

#### <a name="getregexmatches--object"></a>getRegExMatches() → {Object}

選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。

> [!NOTE]
> IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。

`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。

たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|読み取り|

##### <a name="returns"></a>戻り値:

マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。

<dl class="param-type">

<dt>型</dt>

<dd>Object</dd>

</dl>

##### <a name="example"></a>例

次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a>getRegExMatchesByName(name)] → [(許容) {配列。 < 文字列 >}

選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。

> [!NOTE]
> IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。

`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。

アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。

##### <a name="parameters"></a>パラメーター:

|名前|種類|説明|
|---|---|---|
|`name`|String|一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。|

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|読み取り|

##### <a name="returns"></a>戻り値:

マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。

<dl class="param-type">

<dt>型</dt>

<dd>配列。 < 文字列 ></dd>

</dl>

##### <a name="example"></a>例

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a>getSelectedDataAsync(coercionType, [options], callback) → {String}

メッセージの件名または本文から非同期的に選択したデータを返します。

選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。

##### <a name="parameters"></a>パラメーター:

|名前|型|属性|説明|
|---|---|---|---|
|`coercionType`|[Office.CoercionType](office.md#coerciontype-string)||データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。|
|`options`|Object|&lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`|Object|&lt;optional&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`|function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。<br/><br/>コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。 選択範囲は、source プロパティにアクセスするには、呼び出す`asyncResult.value.sourceProperty`、いずれかの方法となる`body`または`subject`。|

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.2|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成|

##### <a name="returns"></a>戻り値:

選択されたデータ (`coercionType` で決定された形式の文字列)。

<dl class="param-type">

<dt>型</dt>

<dd>String</dd>

</dl>

##### <a name="example"></a>例

```
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a>getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}

強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。

> [!NOTE]
> IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.6|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|読み取り|

##### <a name="returns"></a>戻り値:

型:[Entities](/javascript/api/outlook/office.entities)

##### <a name="example"></a>例

次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a>getSelectedRegExMatches() → {Object}

マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。

> [!NOTE]
> IOS は、Outlook または Outlook Android のでは、このメソッドはサポートされていません。

`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。

たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.6|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|読み取り|

##### <a name="returns"></a>戻り値:

マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。

##### <a name="example"></a>例

次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a>getSharedPropertiesAsync ([オプション] では、コールバック)

共有フォルダー、予定表、またはメールボックス内の選択されている予定またはメッセージのプロパティを取得します。

##### <a name="parameters"></a>パラメーター:

|名前|型|属性|説明|
|---|---|---|---|
|`options`|オブジェクト|&lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`|Object|&lt;optional&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`|function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。<br/><br/>共有のプロパティはそのまま、[`SharedProperties`](/javascript/api/outlook/office.sharedproperties)オブジェクトで、`asyncResult.value`プロパティ。 このオブジェクトは、アイテムの共有のプロパティの取得に使用できます。|

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|プレビュー|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成または読み取り|

##### <a name="example"></a>例

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a>loadCustomPropertiesAsync(callback, [userContext])

選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。

カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。

##### <a name="parameters"></a>パラメーター:

|名前|型|属性|説明|
|---|---|---|---|
|`callback`|function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。<br/><br/>カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties) オブジェクトとして指定されます。 取得し、アイテムのカスタム プロパティを削除してサーバーにバックアップを設定するカスタム プロパティに対する変更を保存するのには、このオブジェクトを使用できます。|
|`userContext`|オブジェクト|&lt;省略可能&gt;|開発者は、コールバック関数にアクセスする任意のオブジェクトを提供できます。 によってこのオブジェクトにアクセスできる、`asyncResult.asyncContext`コールバック関数のプロパティです。|

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成または読み取り|

##### <a name="example"></a>例

次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。

```
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a>removeAttachmentAsync(attachmentId, [options], [callback])

メッセージまたは予定から添付ファイルを削除します。

`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。

##### <a name="parameters"></a>パラメーター:

|名前|型|属性|説明|
|---|---|---|---|
|`attachmentId`|String||削除する添付ファイルの識別子。文字列の最大長は 100 文字です。|
|`options`|Object|&lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`|Object|&lt;optional&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`|function|&lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。 <br/>添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。|

##### <a name="errors"></a>エラー

|エラー コード|説明|
|------------|-------------|
|`InvalidAttachmentId`|添付ファイル識別子が存在しません。|

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.1|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成|

##### <a name="example"></a>例

次のコードは、'0' の識別子を持つ添付ファイルを削除します。

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="removehandlerasynceventtype-handler-options-callback"></a>removeHandlerAsync (イベントの種類、ハンドラー、[オプション]、[コールバック])

サポートされているイベントのイベント ハンドラーを削除します。

現在サポートされているイベントの種類は、 `Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged`と`Office.EventType.RecurrenceChanged`

##### <a name="parameters"></a>パラメーター:

| 名前 | 型 | 属性 | 説明 |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || ハンドラーを呼び出す必要のあるイベント。 |
| `handler` | Function || イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`removeHandlerAsync` に渡される `eventType` パラメーターと一致します。 |
| `options` | Object | &lt;optional&gt; | 次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。 |
| `options.asyncContext` | Object | &lt;optional&gt; | 開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。 |
| `callback` | function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.7 |
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り |

####  <a name="saveasyncoptions-callback"></a>saveAsync([options], callback)

項目を非同期的に保存します。

呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。

> [!NOTE]
> アドインを呼び出す場合は、`saveAsync`内のアイテムの作成モードを取得するのには、 `itemId` EWS または REST API を使用するにすると、Outlook キャッシュ モードでは、かかる場合がある項目が実際には、サーバーと同期をとる前にいくつかの時間に注意してください。 使用して、項目が同期されるまで、`itemId`エラーが返されます。

予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。

> [!NOTE]
> 次のクライアントのさまざまな問題のある`saveAsync`の予定の作成モード。
>
> - Mac の Outlook をサポートしていない`saveAsync`での会議では、作成モードです。 呼び出す`saveAsync`Mac の Outlook で会議のエラーが返されます。
> - Web 上で outlook が常に招待状を送信または更新する場合`saveAsync`予定で作成モードです。

##### <a name="parameters"></a>パラメーター:

|名前|型|属性|説明|
|---|---|---|---|
|`options`|オブジェクト|&lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`|Object|&lt;optional&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`|function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。<br/><br/>成功した場合、項目の識別子が提供されている、`asyncResult.value`プロパティ。|

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.3|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成|

##### <a name="examples"></a>例

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a>setSelectedDataAsync(data, [options], callback)

メッセージの本文または件名に非同期的にデータを挿入します。

`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。

##### <a name="parameters"></a>パラメーター:

|名前|型|属性|説明|
|---|---|---|---|
|`data`|String||挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。|
|`options`|Object|&lt;optional&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`|Object|&lt;optional&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`options.coercionType`|[Office.CoercionType](office.md#coerciontype-string)|&lt;optional&gt;|`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。<br/><br/>`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。<br/><br/>`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。|
|`callback`|function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。|

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)|1.2|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)|作成|

##### <a name="example"></a>例

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```