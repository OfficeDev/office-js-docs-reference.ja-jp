 

# <a name="office"></a>Office

Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共有 API](/javascript/api/office)」を参照してください。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="members-and-methods"></a>メンバーとメソッド

| メンバー | 種類 |
|--------|------|
| [AsyncResultStatus](#asyncresultstatus-string) | メンバー |
| [CoercionType](#coerciontype-string) | メンバー |
| [EventType](#eventtype-string) | メンバー |
| [SourceProperty](#sourceproperty-string) | メンバー |

### <a name="namespaces"></a>名前空間

[コンテキスト](office.context.md): Outlook アドイン API で使用するための Office アドイン API のコンテキストの名前空間からの共有インターフェイスを提供します。

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype):ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。

### <a name="members"></a>メンバー

####  <a name="asyncresultstatus-string"></a>AsyncResultStatus :String

非同期呼び出しの結果を指定します。

##### <a name="type"></a>型:

*   String

##### <a name="properties"></a>プロパティ:

|名前| 種類| 説明|
|---|---|---|
|`Succeeded`| String|呼び出しが成功しました。|
|`Failed`| String|呼び出しが失敗しました。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

---

####  <a name="coerciontype-string"></a>CoercionType :String

呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。

##### <a name="type"></a>型:

*   String

##### <a name="properties"></a>プロパティ:

|名前| 種類| 説明|
|---|---|---|
|`Html`| String|HTML 形式で返されるデータを要求します。|
|`Text`| String|テキスト形式で返されるデータを要求します。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

---

####  <a name="eventtype-string"></a>EventType :String

イベント ハンドラーに関連付けられているイベントを指定します。

##### <a name="type"></a>型:

*   String

##### <a name="properties"></a>プロパティ:

| 名前 | 種類 | 説明 | 最小要件セット |
|---|---|---|---|
|`AppointmentTimeChanged`| String | 日付または時刻を選択した予定の系列が変更されました。 | 1.7 |
|`ItemChanged`| String | 選択したアイテムが変更されました。 | 1.5 |
|`OfficeThemeChanged`| String | 選択したアイテムが変更されました。 | Preview |
|`RecipientsChanged`| String | 選択したアイテムまたは予定の場所の受信者の一覧が変更されました。 | 1.7 |
|`RecurrenceChanged`| String | 選択した一連の定期的なパターンが変更されています。 | 1.7 |

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り |

---

####  <a name="sourceproperty-string"></a>SourceProperty :String

呼び出されたメソッドによって返されるデータのソースを指定します。

##### <a name="type"></a>型:

*   String

##### <a name="properties"></a>プロパティ:

|名前| 種類| 説明|
|---|---|---|
|`Body`| String|データのソースは、メッセージの本文です。|
|`Subject`| String|データのソースは、メッセージの件名です。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|