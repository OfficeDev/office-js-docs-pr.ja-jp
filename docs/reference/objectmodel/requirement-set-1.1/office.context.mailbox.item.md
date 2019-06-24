---
title: Office. メールボックス-要件セット1.1
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 09956256f856db6a8bea3053ecb64a2ac2a5fcd9
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128469"
---
# <a name="item"></a>item

### [Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item

`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| 制限あり|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

### <a name="example"></a>例

次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。

```javascript
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
};
```

### <a name="members"></a>メンバー

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a>添付ファイル: <[Attachmentdetails](/javascript/api/outlook_1_1/office.attachmentdetails)>

アイテムの添付ファイルの配列を取得します。閲覧モードのみ。

> [!NOTE]
> セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。 詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。

##### <a name="type"></a>型

*   Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)>

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。

```javascript
var item = Office.context.mailbox.item;
var outputString = "";

if (item.attachments.length > 0) {
  for (i = 0 ; i < item.attachments.length ; i++) {
    var attachment = item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += attachment.name;
    outputString += "<BR>ID: " + attachment.id;
    outputString += "<BR>contentType: " + attachment.contentType;
    outputString += "<BR>size: " + attachment.size;
    outputString += "<BR>attachmentType: " + attachment.attachmentType;
    outputString += "<BR>isInline: " + attachment.isInline;
  }
}

console.log(outputString);
```

#### <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a>bcc:[受信者](/javascript/api/outlook_1_1/office.recipients)

メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。 新規作成モードのみ。

##### <a name="type"></a>型

*   [受信者](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 作成|

##### <a name="example"></a>例

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlook11officebody"></a>本文:[本文](/javascript/api/outlook_1_1/office.body)

アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。

##### <a name="type"></a>型

*   [Body](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

##### <a name="example"></a>例

この例では、メッセージの本文をプレーン テキストで取得します。

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

次の例は、コールバック関数に渡される結果パラメーターの例です。

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a>cc: <[emailaddressdetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_1/office.recipients)

メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。 オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。

##### <a name="read-mode"></a>閲覧モード

`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a>新規作成モード

`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a>型

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

#### <a name="nullable-conversationid-string"></a>(nullable) conversationId: String

特定のメッセージが含まれている電子メールの会話の識別子を取得します。

メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。

新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

##### <a name="example"></a>例

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a>dateTimeCreated: Date

アイテムが作成された日時を取得します。閲覧モードのみ。

##### <a name="type"></a>型

*   日付

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a>dateTimeModified: Date

アイテムが最後に変更された日時を取得します。 Read mode only.

> [!NOTE]
> このメンバーは、iOS または Android の Outlook ではサポートされていません。

##### <a name="type"></a>型

*   日付

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlook11officetime"></a>終了: 日付 |[時間](/javascript/api/outlook_1_1/office.time)

予定が終了する日時を取得または設定します。

`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。

##### <a name="read-mode"></a>閲覧モード

`end` プロパティは `Date` オブジェクトを返します。

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a>新規作成モード

`end` プロパティは `Time` オブジェクトを返します。

[`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。

次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。

```javascript
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a>型

*   Date | [Time](/javascript/api/outlook_1_1/office.time)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a>from: [Emailaddressdetails](/javascript/api/outlook_1_1/office.emailaddressdetails)

メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。

メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。

> [!NOTE]
> `from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。

##### <a name="type"></a>型

*   [EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a>internetMessageId: String

電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a>itemClass: String

選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。

`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。

| 型 | 説明 | アイテム クラス |
| --- | --- | --- |
| 予定アイテム | アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。 | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| メッセージ アイテム | これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。 | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。

##### <a name="type"></a>型

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a>(nullable) itemId: String

現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。 閲覧モードのみ。

> [!NOTE]
> `itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。 `itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。 この値を使用して REST API を呼び出す前に、を`Office.context.mailbox.convertToRestId`使用して変換する必要があります。これは、要件セット1.3 から開始できます。 詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a>itemType: [MailboxEnums](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

インスタンスが表しているアイテムの種類を取得します。

`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。

##### <a name="type"></a>型

*   [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

##### <a name="example"></a>例

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlook11officelocation"></a>場所: String |[場所](/javascript/api/outlook_1_1/office.location)

予定の場所を取得または設定します。

##### <a name="read-mode"></a>閲覧モード

`location` プロパティは、予定の場所を格納した文字列を返します。

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a>新規作成モード

`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a>型

*   String | [Location](/javascript/api/outlook_1_1/office.location)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

#### <a name="normalizedsubject-string"></a>normalizedSubject: String

すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。

normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a>任意出席者: 配列. <[emailaddressdetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_1/office.recipients)

イベントの任意出席者へのアクセスを提供します。 オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。

##### <a name="read-mode"></a>閲覧モード

`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a>新規作成モード

`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a>型

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a>開催者: [Emailaddressdetails](/javascript/api/outlook_1_1/office.emailaddressdetails)

指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。

##### <a name="type"></a>型

*   [EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a>requiredatて dees: 配列. <[emailaddressdetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_1/office.recipients)

イベントの必須出席者へのアクセスを提供します。 オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。

##### <a name="read-mode"></a>閲覧モード

`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a>新規作成モード

`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a>型

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a>sender: [Emailaddressdetails](/javascript/api/outlook_1_1/office.emailaddressdetails)

電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。

メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。

> [!NOTE]
> `from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。

##### <a name="type"></a>型

*   [EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlook11officetime"></a>開始: 日付 |[時間](/javascript/api/outlook_1_1/office.time)

予定を開始する日時を取得または設定します。

`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。

##### <a name="read-mode"></a>閲覧モード

`start` プロパティは `Date` オブジェクトを返します。

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a>新規作成モード

`start` プロパティは `Time` オブジェクトを返します。

[`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。

次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。

```javascript
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a>型

*   Date | [Time](/javascript/api/outlook_1_1/office.time)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

#### <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a>subject: String |[件名](/javascript/api/outlook_1_1/office.subject)

アイテムの件名フィールドに示される説明を取得または設定します。

`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。

##### <a name="read-mode"></a>閲覧モード

`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a>新規作成モード

`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a>型

*   String | [Subject](/javascript/api/outlook_1_1/office.subject)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a>宛先: 配列. <[emailaddressdetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_1/office.recipients)

メッセージの **To** 行にある受信者へのアクセスを提供します。 オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。

##### <a name="read-mode"></a>閲覧モード

`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a>新規作成モード

`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a>型

*   Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

### <a name="methods"></a>メソッド

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a>addFileAttachmentAsync(uri, attachmentName, [options], [callback])

ファイルを添付ファイルとしてメッセージまたは予定に追加します。

`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。

その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。

##### <a name="parameters"></a>パラメーター

|名前| 種類| 属性| 説明|
|---|---|---|---|
|`uri`| String||メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。|
|`attachmentName`| String||添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。|
|`options`| オブジェクト| &lt;オプション&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`| オブジェクト| &lt;省略可能&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。 <br/>成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。<br/>添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。|

##### <a name="errors"></a>エラー

| エラー コード | 説明 |
|------------|-------------|
| `AttachmentSizeExceeded` | 添付ファイルのサイズが上限を超えています。 |
| `FileTypeNotSupported` | 許可されていない拡張子の添付ファイルです。 |
| `NumberOfAttachmentsExceeded` | メッセージまたは予定の添付ファイルが多すぎます。 |

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 作成|

##### <a name="example"></a>例

```javascript
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a>addItemAttachmentAsync(itemId, attachmentName, [options], [callback])

メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。

`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。

その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。

Office アドインが Outlook on the web で実行されている場合、 `addItemAttachmentAsync`メソッドは、編集しているアイテム以外のアイテムにアイテムを添付できます。ただし、これはサポートされておらず、推奨されていません。

##### <a name="parameters"></a>パラメーター

|名前| 型| 属性| 説明|
|---|---|---|---|
|`itemId`| String||添付するアイテムの Exchange 識別子。最大長は 100 文字です。|
|`attachmentName`| String||添付するアイテムの件名。 最大の長さは、255 文字です。|
|`options`| Object| &lt;オプション&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`| オブジェクト| &lt;省略可能&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`| function| &lt;任意&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。 <br/>成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。<br/>添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。|

##### <a name="errors"></a>エラー

| エラー コード | 説明 |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | メッセージまたは予定の添付ファイルが多すぎます。 |

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 作成|

##### <a name="example"></a>例

次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。

```javascript
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach (shortened for readability).
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

#### <a name="displayreplyallformformdata-callback"></a>displayReplyAllForm(formData, [callback])

選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。

> [!NOTE]
> このメソッドは、iOS または Android の Outlook ではサポートされていません。

Web 上の Outlook では、返信フォームは、3列表示のポップアップフォームとして表示され、2列または1列表示のポップアップフォームとして表示されます。

文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。

> [!NOTE]
> へ`displayReplyAllForm`の呼び出しに添付ファイルを含める機能は、要件セット1.1 ではサポートされていません。 添付ファイルのサポートは、要件セット 1.2 以降で `displayReplyAllForm` に追加されました。

##### <a name="parameters"></a>パラメーター

|名前| 型| 説明|
|---|---|---|
|`formData`| String &#124; Object| |回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。<br/>**または**<br/>本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。 |
| `formData.htmlBody` | String | &lt;省略可能&gt; | 回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。
| `callback` | function | &lt;optional&gt; | メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。 |

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="examples"></a>例

次のコードは `displayReplyAllForm` 関数に文字列を渡します。

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

空の本文を返信します。

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

本文だけを返信します。

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

本文とコールバックを返信します。

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata-callback"></a>displayReplyForm(formData, [callback])

選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。

> [!NOTE]
> このメソッドは、iOS または Android の Outlook ではサポートされていません。

Web 上の Outlook では、返信フォームは、3列表示のポップアップフォームとして表示され、2列または1列表示のポップアップフォームとして表示されます。

文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。

> [!NOTE]
> へ`displayReplyForm`の呼び出しに添付ファイルを含める機能は、要件セット1.1 ではサポートされていません。 添付ファイルのサポートは、要件セット 1.2 以降で `displayReplyForm` に追加されました。

##### <a name="parameters"></a>パラメーター

|名前| 型| 説明|
|---|---|---|
|`formData`| String &#124; Object| | 回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。<br/>**または**<br/>本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。 |
| `formData.htmlBody` | String | &lt;省略可能&gt; | 回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。
| `callback` | function | &lt;optional&gt; | メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。 |

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="examples"></a>例

次のコードは `displayReplyForm` 関数に文字列を渡します。

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

空の本文を返信します。

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

本文だけを返信します。

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

本文とコールバックを返信します。

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a>getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}

選択したアイテムの本文にあるエンティティを取得します。

> [!NOTE]
> このメソッドは、iOS または Android の Outlook ではサポートされていません。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="returns"></a>戻り値:

型:[Entities](/javascript/api/outlook_1_1/office.entities)

##### <a name="example"></a>例

次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a>getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}

選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。

> [!NOTE]
> このメソッドは、iOS または Android の Outlook ではサポートされていません。

##### <a name="parameters"></a>パラメーター

|名前| 型| 説明|
|---|---|---|
|`entityType`| [Office.MailboxEnums.EntityType](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|EntityType 列挙値の 1 つ。|

##### <a name="requirements"></a>Requirements

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| 制限あり|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="returns"></a>戻り値:

`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。 指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。 それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。

このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。

| `entityType` の値 | 返される配列内のオブジェクトの型 | 必要なアクセス許可のレベル |
| --- | --- | --- |
| `Address` | 文字列 | **制限あり** |
| `Contact` | 連絡先 | **ReadItem** |
| `EmailAddress` | 文字列 | **ReadItem** |
| `MeetingSuggestion` | MeetingSuggestion | **ReadItem** |
| `PhoneNumber` | PhoneNumber | **制限あり** |
| `TaskSuggestion` | TaskSuggestion | **ReadItem** |
| `URL` | 文字列 | **制限あり** |

型:Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>


##### <a name="example"></a>例

次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。

```javascript
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a>getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}

マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。

> [!NOTE]
> このメソッドは、iOS または Android の Outlook ではサポートされていません。

`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。

##### <a name="parameters"></a>パラメーター

|名前| 型| 説明|
|---|---|---|
|`name`| String|一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="returns"></a>戻り値:

`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。


型:Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>


#### <a name="getregexmatches--object"></a>getRegExMatches() → {Object}

選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。

> [!NOTE]
> このメソッドは、iOS または Android の Outlook ではサポートされていません。

`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。

たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="returns"></a>戻り値:

マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。

<dl class="param-type">

<dt>型</dt>

<dd>Object</dd>

</dl>

##### <a name="example"></a>例

次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a>getRegExMatchesByName(name) → (nullable) {Array.< String >}

選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。

> [!NOTE]
> このメソッドは、iOS または Android の Outlook ではサポートされていません。

`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。

アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。

##### <a name="parameters"></a>パラメーター

|名前| 型| 説明|
|---|---|---|
|`name`| String|一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="returns"></a>戻り値:

マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。

<dl class="param-type">

<dt>型</dt>

<dd>Array.< String ></dd>

</dl>

##### <a name="example"></a>例

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="loadcustompropertiesasynccallback-usercontext"></a>loadCustomPropertiesAsync(callback, [userContext])

選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。

カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。

##### <a name="parameters"></a>パラメーター

|名前| 型| 属性| 説明|
|---|---|---|---|
|`callback`| function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。<br/><br/>カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) オブジェクトとして指定されます。 このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。|
|`userContext`| オブジェクト| &lt;省略可能&gt;|開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。 このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

##### <a name="example"></a>例

次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。

```javascript
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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a>removeAttachmentAsync(attachmentId, [options], [callback])

メッセージまたは予定から添付ファイルを削除します。

`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。 ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。 Outlook on the web およびモバイルデバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。 ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。

##### <a name="parameters"></a>パラメーター

|名前| 型| 属性| 説明|
|---|---|---|---|
|`attachmentId`| String||削除する添付ファイルの識別子。|
|`options`| オブジェクト| &lt;オプション&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`| オブジェクト| &lt;省略可能&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`| function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。 <br/>添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。|

##### <a name="errors"></a>エラー

| エラー コード | 説明 |
|------------|-------------|
| `InvalidAttachmentId` | 添付ファイル識別子が存在しません。 |

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 作成|

##### <a name="example"></a>例

次のコードは、'0' の識別子を持つ添付ファイルを削除します。

```javascript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```
