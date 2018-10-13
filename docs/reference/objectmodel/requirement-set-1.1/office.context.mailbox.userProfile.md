
# <a name="userprofile"></a>userProfile

### [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックス要件セットの最小バージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または閲覧|

### <a name="members"></a>メンバー

####  <a name="displayname-string"></a>displayName : 文字列

ユーザーの表示名を取得します。

##### <a name="type"></a>種類:

*   文字列

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックス要件セットの最小バージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または閲覧|

##### <a name="example"></a>例

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a>emailAddress : 文字列

ユーザーの SMTP 電子メール アドレスを取得します。

##### <a name="type"></a>型:

*   文字列

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックス要件セットの最小バージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または閲覧|

##### <a name="example"></a>例

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a>タイム ゾーン : 文字列

ユーザーの既定のタイム ゾーンを取得します。

##### <a name="type"></a>種類:

*   文字列

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックス要件セットの最小バージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または閲覧|

##### <a name="example"></a>例

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```