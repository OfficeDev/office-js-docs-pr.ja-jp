---
title: Office. メールボックス要件セット1.2
description: ''
ms.date: 10/21/2019
localization_priority: Normal
ms.openlocfilehash: 542e8c9899c2d4a3c5b4546c3d5a73ba0d3c3a7e
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/22/2019
ms.locfileid: "37627000"
---
# <a name="mailbox"></a>mailbox

### <a name="officeofficemdcontextofficecontextmdmailbox"></a>[Office](Office.md)[.context](Office.context.md).mailbox

Microsoft Outlook の Outlook アドイン オブジェクト モデルへのアクセスを提供します。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| 制限あり|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

##### <a name="members-and-methods"></a>メンバーとメソッド

| メンバー | 種類 |
|--------|------|
| [ewsUrl](#ewsurl-string) | メンバー |
| [convertToLocalClientTime](#converttolocalclienttimetimevalue--localclienttime) | メソッド |
| [convertToUtcClientTime](#converttoutcclienttimeinput--date) | メソッド |
| [displayAppointmentForm](#displayappointmentformitemid) | メソッド |
| [displayMessageForm](#displaymessageformitemid) | メソッド |
| [displayNewAppointmentForm](#displaynewappointmentformparameters) | メソッド |
| [getCallbackTokenAsync](#getcallbacktokenasynccallback-usercontext) | メソッド |
| [getUserIdentityTokenAsync](#getuseridentitytokenasynccallback-usercontext) | メソッド |
| [makeEwsRequestAsync](#makeewsrequestasyncdata-callback-usercontext) | メソッド |

### <a name="namespaces"></a>名前空間

[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。

[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。

[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。

### <a name="members"></a>Members

#### <a name="ewsurl-string"></a>ewsUrl: String

このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。読み取りモードのみです。

> [!NOTE]
> このメンバーは、Outlook on iOS または Android ではサポートされていません。


  `ewsUrl` 値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。

##### <a name="type"></a>型

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 読み取り|

### <a name="methods"></a>メソッド

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-12"></a>convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)}

クライアントのローカル時間で時間情報が含まれている辞書を取得します。

Outlook on the web または Outlook デスクトップのメールアプリでは、日付と時刻に異なるタイムゾーンを使用できます。Outlook デスクトップは、クライアント コンピューターのタイム ゾーンを使用します。Outlook on the web は、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使用します。日付と時刻の値は、ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように処理する必要があります。

Outlook デスクトップ クライアントでメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook on the web でメール アプリを実行している場合、`convertToLocalClientTime` メソッドは、EAC で指定したタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。

##### <a name="parameters"></a>パラメーター

|名前| 種類| 説明|
|---|---|---|
|`timeValue`| 日付|日付オブジェクト|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

##### <a name="returns"></a>戻り値:

型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a>convertToUtcClientTime(input) → {Date}

時間情報が含まれているディクショナリから日付オブジェクトを取得します。

`convertToUtcClientTime` メソッドは、ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。

##### <a name="parameters"></a>パラメーター

|名前| 種類| 説明|
|---|---|---|
|`input`| [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)|変換するローカル時刻の値。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

##### <a name="returns"></a>戻り値:

時間が UTC で表現された日付オブジェクト。

型: Date

##### <a name="example"></a>例

```js
// Represents 3:37 PM PDT on Monday, August 26, 2019.
var input = {
  date: 26,
  hours: 15,
  milliseconds: 2,
  minutes: 37,
  month: 7,
  seconds: 2,
  timezoneOffset: -420,
  year: 2019
};

// result should be a Date object.
var result = Office.context.mailbox.convertToUtcClientTime(input);

// Output should be "2019-08-26T22:37:02.002Z".
console.log(result.toISOString());
```

<br>

---
---

#### <a name="displayappointmentformitemid"></a>displayAppointmentForm(itemId)

既存の予定を表示します。

> [!NOTE]
> このメソッドは、Outlook on iOS または Android ではサポートされていません。

`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存の予定を開きます。

Outlook on Mac では、このメソッドを使用して定期的な系列に含まれない単発の予定や定期的な系列のマスター予定を表示できます。ただし、系列のインスタンスは表示できません。これは、Outlook on Mac では定期的な系列のインスタンスのプロパティ (アイテム ID を含む) にアクセスできないためです。

Outlook on the web では、このメソッドはフォームの本文が 32KB 以下の文字数の場合にのみ指定のフォームを開きます。

指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピューターまたはデバイスで空のウィンドウが開き、エラー メッセージは返されません。

##### <a name="parameters"></a>パラメーター

|名前| 種類| 説明|
|---|---|---|
|`itemId`| String|既存の予定の Exchange Web サービス (EWS) 識別子。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

##### <a name="example"></a>例

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a>displayMessageForm(itemId)

既存のメッセージを表示します。

> [!NOTE]
> このメソッドは、Outlook on iOS または Android ではサポートされていません。

`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスに既存のメッセージを開きます。

Outlook on the web では、このメソッドはフォームの本文が 32 KB 以下の文字数の場合にのみ指定のフォームを開きます。

指定のアイテム識別子が既存のメッセージを表していない場合は、クライアント コンピューターにはメッセージは表示されず、エラー メッセージも返されません。

予定を表す `itemId` を含む `displayMessageForm` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用します。新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。

##### <a name="parameters"></a>パラメーター

|名前| 種類| Description|
|---|---|---|
|`itemId`| String|既存のメッセージの Exchange Web サービス (EWS) 識別子。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

##### <a name="example"></a>例

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a>displayNewAppointmentForm(parameters)

新しい予定を作成するためのフォームを表示します。

> [!NOTE]
> このメソッドは、Outlook on iOS または Android ではサポートされていません。

`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメーターを指定すると、予定のフォーム フィールドにパラメーターの内容が自動的に設定されます。

Outlook on the web およびモバイル デバイスでは、このメソッドは常に出席者フィールドが含まれるフォームを表示します。入力引数として出席者を指定しないと、このメソッドは **[保存]** ボタンのあるフォームを表示します。出席者を指定した場合には、フォームにその出席者と **[送信]** ボタンが表示されます。

Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメーターに出席者またはリソースを指定し、このメソッドを実行すると、**[送信]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[保存して閉じる]** ボタンがある予定フォームが表示されます。

パラメーターのいずれかが指定のサイズ制限を超える場合、または不明なパラメーター名が指定されている場合は、例外がスローされます。

##### <a name="parameters"></a>パラメーター

|名前| 種類| 説明|
|---|---|---|
| `parameters` | Object | 新しい予定を記述するパラメーターのディクショナリ。 |
| `parameters.requiredAttendees` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt; | 予定に必要な各出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。 |
| `parameters.optionalAttendees` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt; | 予定の各任意出席者について、メール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。 |
| `parameters.start` | 日付 | 予定の開始日時を指定する `Date` オブジェクト。 |
| `parameters.end` | 日付 | 予定の終了日時を指定する `Date` オブジェクト。 |
| `parameters.location` | String | 予定の場所を含む文字列。文字列は最大 255 文字に制限されます。 |
| `parameters.resources` | Array.&lt;String&gt; | 予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。 |
| `parameters.subject` | String | 予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。 |
| `parameters.body` | String | 予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。 |

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```js
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

<br>

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a>getCallbackTokenAsync(callback, [userContext])

Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。

`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。

トークンと、添付ファイル識別子またはアイテム識別子の両方をサードパーティのシステムに渡すことができます。 サードパーティシステムは、トークンをベアラー認証トークンとして使用して、Exchange Web サービス (EWS) の[Getattachment](/exchange/client-developer/web-service-reference/getattachment-operation)操作または[GetItem](/exchange/client-developer/web-service-reference/getitem-operation)操作を呼び出して、添付ファイルまたはアイテムを返します。 For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).

メソッドを`getCallbackTokenAsync`呼び出すには、 **ReadItem**の最低限のアクセス許可レベルが必要です。

##### <a name="parameters"></a>パラメーター

|名前| 種類| 属性| 説明|
|---|---|---|---|
|`callback`| function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。<br/><br/>トークンは、`asyncResult.value` プロパティで文字列として提供されます。<br><br>エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。|
|`userContext`| オブジェクト| &lt;省略可能&gt;|非同期メソッドに渡される状態データです。|

##### <a name="errors"></a>エラー

|エラー コード|説明|
|------------|-------------|
|`HTTPRequestFailure`|要求が失敗しました。 HTTP エラーコードの diagnostics オブジェクトを参照してください。|
|`InternalServerError`|Exchange サーバーがエラーを返しました。 詳細については、diagnostics オブジェクトを参照してください。|
|`NetworkError`|ユーザーはネットワークに接続されていません。 ネットワーク接続を確認し、やり直してください。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getuseridentitytokenasynccallback-usercontext"></a>getUserIdentityTokenAsync(callback, [userContext])

ユーザーと Office アドインを識別するトークンを取得します。

`getUserIdentityTokenAsync` メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](/outlook/add-ins/authentication)することのできるトークンを返します。

##### <a name="parameters"></a>パラメーター

|名前| 種類| 属性| 説明|
|---|---|---|---|
|`callback`| function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。<br/><br/>トークンは、`asyncResult.value` プロパティで文字列として提供されます。<br><br>エラーが発生した場合、 `asyncResult.error` および `asyncResult.diagnostics` のプロパティで追加情報が提供される場合があります。|
|`userContext`| オブジェクト| &lt;省略可能&gt;|非同期メソッドに渡される状態データです。|

##### <a name="errors"></a>エラー

|エラー コード|説明|
|------------|-------------|
|`HTTPRequestFailure`|要求が失敗しました。 HTTP エラーコードの diagnostics オブジェクトを参照してください。|
|`InternalServerError`|Exchange サーバーがエラーを返しました。 詳細については、diagnostics オブジェクトを参照してください。|
|`NetworkError`|ユーザーはネットワークに接続されていません。 ネットワーク接続を確認し、やり直してください。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

##### <a name="example"></a>例

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a>makeEwsRequestAsync(data, callback, [userContext])

ユーザーのメールボックスをホストしている Exchange サーバー上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。

> [!NOTE]
> このメソッドは、次のシナリオではサポートされていません。
> - Outlook on iOS または Android の場合
> - アドインが Gmail のメールボックスに読み込まれる場合
> 
> このような場合は、アドインでは [REST API を使用](/outlook/add-ins/use-rest-api)して、代わりにユーザーのメールボックスにアクセスする必要があります。

The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange. See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.

`makeEwsRequestAsync` メソッドでは、フォルダー関連アイテムを要求できません。

XML 要求では UTF-8 エンコードを指定する必要があります。

```xml
<?xml version="1.0" encoding="utf-8"?>
```

`makeEwsRequestAsync` メソッドを使用するには、アドインに **ReadWriteMailbox** アクセス許可が必要です。**ReadWriteMailbox** アクセス許可と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作の使い方については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。

> [!NOTE]
> サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。

##### <a name="version-differences"></a>バージョンの相違点

バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

Outlook on the web でメール アプリを実行している場合は、エンコード値を設定する必要はありません。mailbox.diagnostics.hostName プロパティを使って、メール アプリを Outlook で実行しているのか、Outlook on the web で実行しているのかを確認できます。mailbox.diagnostics.hostVersion プロパティを使って、どのバージョンの Outlook を使って実行しているかを確認できます。

##### <a name="parameters"></a>パラメーター

|名前| 種類| 属性| 説明|
|---|---|---|---|
|`data`| String||EWS 要求です。|
|`callback`| function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。<br/><br/>The XML result of the EWS call is provided as a string in the `asyncResult.value` property. If the result exceeds 1 MB in size, an error message is returned instead.|
|`userContext`| オブジェクト| &lt;省略可能&gt;|非同期メソッドに渡される状態データです。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteMailbox|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

##### <a name="example"></a>例

次の例は、`makeEwsRequestAsync` を呼び出し、`GetItem` 操作を使って項目の件名を取得します。

```js
function getSubjectRequest(id) {
  // Return a GetItem operation request for the subject of the specified item.
  var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

  return request;
}

function sendRequest() {
  // Create a local variable that contains the mailbox.
  Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
  var result = asyncResult.value;
  var context = asyncResult.asyncContext;

  // Process the returned response here.
}
```
