
# <a name="userprofile"></a>userProfile

### [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="members-and-methods"></a>メンバーとメソッド

| メンバー | 種類 |
|--------|------|
| [accountType](#accounttype-string) | Member |
| [displayName](#displayname-string) | メンバー |
| [emailAddress](#emailaddress-string) | メンバー |
| [timeZone](#timezone-string) | メンバー |

### <a name="members"></a>Members

####  <a name="accounttype-string"></a>accountType: 文字列

> [!NOTE]
> このメンバーは、現在のみ 2016 の Outlook でサポートされている Mac の後で (ビルド 16.9.1212 またはそれ以降)。

メールボックスに関連付けられているユーザーのアカウントの種類を取得します。 使用可能な値は、次の表に表示されます。

| 値 | 説明 |
|-------|-------------|
| `enterprise` | メールボックスは、オンプレミスの Exchange サーバーには。 |
| `gmail` | メールボックスは、Gmail アカウントに関連付けられます。 |
| `office365` | メールボックスが関連付けられている、Office 365 の機能や、学校のアカウントです。 |
| `outlookCom` | メールボックスは、個人、Outlook.com アカウントに関連付けられます。 |

##### <a name="type"></a>型:

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="example"></a>例

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a>displayName :String

ユーザーの表示名を取得します。

##### <a name="type"></a>型:

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="example"></a>例

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a>emailAddress :String

ユーザーの SMTP 電子メール アドレスを取得します。

##### <a name="type"></a>型:

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="example"></a>例

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a>timeZone :String

ユーザーの既定のタイム ゾーンを取得します。

##### <a name="type"></a>型:

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="example"></a>例

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```