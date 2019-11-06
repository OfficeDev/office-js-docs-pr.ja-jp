---
title: Office. メールボックス-要件セット1.8
description: ''
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 065ea3c74580555c0df1af7b495127a25493b612
ms.sourcegitcommit: 21aa084875c9e07a300b3bbe8852b3e5dd163e1d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/06/2019
ms.locfileid: "38001573"
---
# <a name="item"></a>item

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|制限あり|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|新規作成または閲覧|

##### <a name="members-and-methods"></a>メンバーとメソッド

| メンバー | 種類 |
|--------|------|
| [attachments](#attachments-arrayattachmentdetails) | メンバー |
| [bcc](#bcc-recipients) | メンバー |
| [body](#body-body) | メンバー |
| [categories](#categories-categories) | メンバー |
| [cc](#cc-arrayemailaddressdetailsrecipients) | メンバー |
| [conversationId](#nullable-conversationid-string) | メンバー |
| [dateTimeCreated](#datetimecreated-date) | メンバー |
| [dateTimeModified](#datetimemodified-date) | メンバー |
| [end](#end-datetime) | メンバー |
| [enhancedLocation](#enhancedlocation-enhancedlocation) | メンバー |
| [from](#from-emailaddressdetailsfrom) | メンバー |
| [internetHeaders](#internetheaders-internetheaders) | メンバー |
| [internetMessageId](#internetmessageid-string) | メンバー |
| [itemClass](#itemclass-string) | メンバー |
| [itemId](#nullable-itemid-string) | メンバー |
| [itemType](#itemtype-officemailboxenumsitemtype) | メンバー |
| [location](#location-stringlocation) | メンバー |
| [normalizedSubject](#normalizedsubject-string) | メンバー |
| [notificationMessages](#notificationmessages-notificationmessages) | Member |
| [optionalAttendees](#optionalattendees-arrayemailaddressdetailsrecipients) | メンバー |
| [organizer](#organizer-emailaddressdetailsorganizer) | メンバー |
| [recurrence](#nullable-recurrence-recurrence) | メンバー |
| [requiredAttendees](#requiredattendees-arrayemailaddressdetailsrecipients) | メンバー |
| [sender](#sender-emailaddressdetails) | Member |
| [系列 Id](#nullable-seriesid-string) | Member |
| [start](#start-datetime) | Member |
| [subject](#subject-stringsubject) | メンバー |
| [to](#to-arrayemailaddressdetailsrecipients) | メンバー |
| [addFileAttachmentAsync](#addfileattachmentasyncuri-attachmentname-options-callback) | メソッド |
| [addFileAttachmentFromBase64Async](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | メソッド |
| [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) | メソッド |
| [addItemAttachmentAsync](#additemattachmentasyncitemid-attachmentname-options-callback) | メソッド |
| [close](#close) | メソッド |
| [displayReplyAllForm](#displayreplyallformformdata-callback) | メソッド |
| [displayReplyForm](#displayreplyformformdata-callback) | メソッド |
| [getAllInternetHeadersAsync](#getallinternetheadersasyncoptions-callback) | メソッド |
| [getAttachmentContentAsync](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | メソッド |
| [getAttachmentsAsync](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | メソッド |
| [getEntities](#getentities--entities) | メソッド |
| [getEntitiesByType](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | メソッド |
| [getFilteredEntitiesByName](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | メソッド |
| [getItemIdAsync](#getitemidasyncoptions-callback) | メソッド |
| [getRegExMatches](#getregexmatches--object) | メソッド |
| [getRegExMatchesByName](#getregexmatchesbynamename--nullable-array-string-) | メソッド |
| [getSelectedDataAsync](#getselecteddataasynccoerciontype-options-callback--string) | メソッド |
| [Office.context.mailbox.item.getselectedentities](#getselectedentities--entities) | メソッド |
| [Office.context.mailbox.item.getselectedregexmatches](#getselectedregexmatches--object) | メソッド |
| [getSharedPropertiesAsync](#getsharedpropertiesasyncoptions-callback) | メソッド |
| [loadCustomPropertiesAsync](#loadcustompropertiesasynccallback-usercontext) | メソッド |
| [removeAttachmentAsync](#removeattachmentasyncattachmentid-options-callback) | メソッド |
| [removeHandlerAsync](#removehandlerasynceventtype-options-callback) | メソッド |
| [saveAsync](#saveasyncoptions-callback) | メソッド |
| [setSelectedDataAsync](#setselecteddataasyncdata-options-callback) | メソッド |

### <a name="example"></a>例

次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。

```js
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

### <a name="members"></a>Members

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-18"></a>attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)>

アイテムの添付ファイルを配列として取得します。 閲覧モードのみ。

> [!NOTE]
> セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。 詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。

##### <a name="type"></a>型

*   Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)>

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|読み取り|

##### <a name="example"></a>例

次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。

```js
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

<br>

---
---

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a>bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)

メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。 新規作成モードのみ。

既定では、コレクションは最大 100 人のメンバーに制限されています。 ただし、Windows および Mac では、次の制限が適用されます。

- 最大 500 人のメンバーを取得します。
- 呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。

##### <a name="type"></a>型

*   [受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.1|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|作成|

##### <a name="example"></a>例

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

<br>

---
---

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-18"></a>body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.8)

アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。

##### <a name="type"></a>型

*   [Body](/javascript/api/outlook/office.body?view=outlook-js-1.8)

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.1|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|新規作成または閲覧|

##### <a name="example"></a>例

この例では、メッセージの本文をプレーン テキストで取得します。

```js
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

<br>

---
---

#### <a name="categories-categoriesjavascriptapioutlookofficecategoriesviewoutlook-js-18"></a>カテゴリ:[カテゴリ](/javascript/api/outlook/office.categories?view=outlook-js-1.8)

アイテムのカテゴリを管理するためのメソッドを提供するオブジェクトを取得します。

> [!NOTE]
> このメンバーは、Outlook on iOS または Android ではサポートされていません。

##### <a name="type"></a>種類

*   [Categories](/javascript/api/outlook/office.categories?view=outlook-js-1.8)

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.8|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|新規作成または閲覧|

##### <a name="example"></a>例

この例では、アイテムのカテゴリを取得します。

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a>cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)

メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。 オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。

##### <a name="read-mode"></a>閲覧モード

`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。 既定では、コレクションは最大 100 人のメンバーに制限されています。 ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a>新規作成モード

`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。 既定では、コレクションは最大 100 人のメンバーに制限されています。 ただし、Windows および Mac では、次の制限が適用されます。

- 最大 500 人のメンバーを取得します。
- 呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a>型

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|新規作成または閲覧|

<br>

---
---

#### <a name="nullable-conversationid-string"></a>(nullable) conversationId: String

特定のメッセージが含まれている電子メールの会話の識別子を取得します。

メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。

新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|新規作成または閲覧|

##### <a name="example"></a>例

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a>dateTimeCreated: Date

アイテムが作成された日時を取得します。閲覧モードのみ。

##### <a name="type"></a>型

*   日付

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|読み取り|

##### <a name="example"></a>例

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a>dateTimeModified: Date

アイテムが最後に変更された日時を取得します。閲覧モードのみ。

> [!NOTE]
> このメンバーは、Outlook on iOS または Android ではサポートされていません。

##### <a name="type"></a>種類

*   日付

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|読み取り|

##### <a name="example"></a>例

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-18"></a>end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)

予定が終了する日時を取得または設定します。

`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。

##### <a name="read-mode"></a>閲覧モード

`end` プロパティは `Date` オブジェクトを返します。

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a>新規作成モード

`end` プロパティは `Time` オブジェクトを返します。

[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。

次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。

```js
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

*   Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)

##### <a name="requirements"></a>要件

|必要条件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|新規作成または閲覧|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocationviewoutlook-js-18"></a>enhancedLocation: [enhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)

予定の場所を取得または設定します。

##### <a name="read-mode"></a>閲覧モード

この`enhancedLocation`プロパティは、予定に関連付けられている場所 ( [locationdetails](/javascript/api/outlook/office.locationdetails?view=outlook-js-1.8)オブジェクトで表される) のセットを取得できる[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)オブジェクトを返します。

##### <a name="compose-mode"></a>新規作成モード

この`enhancedLocation`プロパティは、予定の場所を取得、削除、または追加するためのメソッドを提供する[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)オブジェクトを返します。

##### <a name="type"></a>型

*   [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)

##### <a name="requirements"></a>Requirements

|必要条件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.8|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|新規作成または閲覧|

##### <a name="example"></a>例

次の例では、予定に関連付けられている現在の場所を取得します。

```js
Office.context.mailbox.item.enhancedLocation.getAsync(callbackFunction);

function callbackFunction(asyncResult) {
  asyncResult.value.forEach(function (place) {
    console.log("Display name: " + place.displayName);
    console.log("Type: " + place.locationIdentifier.type);
    if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
      console.log("Email address: " + place.emailAddress);
    }
  });
}
```

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18fromjavascriptapioutlookofficefromviewoutlook-js-18"></a>from: [emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[from](/javascript/api/outlook/office.from?view=outlook-js-1.8)

メッセージの送信者の電子メール アドレスを取得します。

メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。

> [!NOTE]
> `from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。

##### <a name="read-mode"></a>閲覧モード

プロパティ`from`は`EmailAddressDetails`オブジェクトを返します。

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a>新規作成モード

プロパティ`from`は、from `From`値を取得するメソッドを提供するオブジェクトを返します。

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a>型

*   [電子メールアドレス](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [の](/javascript/api/outlook/office.from?view=outlook-js-1.8)詳細

##### <a name="requirements"></a>Requirements

|要件|||
|---|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|1.7|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|ReadWriteItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|Read|Compose|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheadersviewoutlook-js-18"></a>internetHeaders: [internetHeaders](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8)

メッセージのカスタムインターネットヘッダーを取得または設定します。 新規作成モードのみ。

##### <a name="type"></a>型

*   [InternetHeaders](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8)

##### <a name="requirements"></a>Requirements

|必要条件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.8|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|作成|

##### <a name="example"></a>例

```js
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

<br>

---
---

#### <a name="internetmessageid-string"></a>internetMessageId: String

電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>要件

|必要条件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|読み取り|

##### <a name="example"></a>例

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a>itemClass: String

選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。

`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。

|型|説明|アイテム クラス|
|---|---|---|
|予定アイテム|アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|メッセージ アイテム|これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|読み取り|

##### <a name="example"></a>例

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a>(nullable) itemId: String

現在のアイテムの[Exchange Web サービスのアイテム識別子](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)を取得します。 閲覧モードのみ。

> [!NOTE]
> `itemId`プロパティによって返される識別子は、 [Exchange Web サービスのアイテム識別子](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)と同じです。 `itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。 この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。 詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。

新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|読み取り|

##### <a name="example"></a>例

次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

<br>

---
---

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-18"></a>itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.8)

インスタンスが表しているアイテムの種類を取得します。

`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。

##### <a name="type"></a>型

*   [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.8)

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|新規作成または閲覧|

##### <a name="example"></a>例

```js
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

<br>

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-18"></a>location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)

予定の場所を取得または設定します。

##### <a name="read-mode"></a>閲覧モード

`location` プロパティは、予定の場所を格納した文字列を返します。

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a>新規作成モード

`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a>型

*   String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|新規作成または閲覧|

<br>

---
---

#### <a name="normalizedsubject-string"></a>normalizedSubject: String

すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。

normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。

##### <a name="type"></a>型

*   String

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|読み取り|

##### <a name="example"></a>例

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-18"></a>notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.8)

アイテムの通知メッセージを取得します。

##### <a name="type"></a>型

*   [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.8)

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.3|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|新規作成または閲覧|

##### <a name="example"></a>例

```js
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a>optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)

イベントの任意出席者へのアクセスを提供します。 オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。

##### <a name="read-mode"></a>閲覧モード

`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。 既定では、コレクションは最大 100 人のメンバーに制限されています。 ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a>新規作成モード

`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。 既定では、コレクションは最大 100 人のメンバーに制限されています。 ただし、Windows および Mac では、次の制限が適用されます。

- 最大 500 人のメンバーを取得します。
- 呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a>型

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)

##### <a name="requirements"></a>Requirements

|必要条件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|新規作成または閲覧|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18organizerjavascriptapioutlookofficeorganizerviewoutlook-js-18"></a>開催者: [emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[開催者](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)

指定した会議の開催者の電子メールアドレスを取得します。

##### <a name="read-mode"></a>閲覧モード

プロパティ`organizer`は、会議の開催者を表す[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)オブジェクトを返します。

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a>新規作成モード

プロパティ`organizer`は、開催者の値を取得するためのメソッドを提供する[オーガナイザー](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)オブジェクトを返します。

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a>型

*   [Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [開催者](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)

##### <a name="requirements"></a>Requirements

|要件|||
|---|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|1.7|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|ReadWriteItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|Read|Compose|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrenceviewoutlook-js-18"></a>(nullable) 定期的なスケジュール:[定期的](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)なアイテム

予定の定期的なパターンを取得または設定します。 会議出席依頼の定期的なパターンを取得します。 予定アイテムの読み取りおよび作成モード。 会議出席依頼アイテムの閲覧モード。

この`recurrence`プロパティは、アイテムが series または series 内のインスタンスの場合、定期的な予定または会議出席依頼に対して[定期的](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)なオブジェクトを返します。 `null`は、単一の予定および1つの予定の会議出席依頼に対して返されます。 `undefined`は、会議出席依頼ではないメッセージに対して返されます。

> 注: 会議出席依頼に`itemClass`は、IPM という値があります。出席依頼。

> 注: 定期的なオブジェクトが`null`の場合は、そのオブジェクトが単一の予定または1つの予定の会議出席依頼であり、データ系列の一部ではないことを示します。

##### <a name="read-mode"></a>閲覧モード

この`recurrence`プロパティは、定期的な予定を表す[定期的](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)オブジェクトを返します。 これは、予定および会議出席依頼に対して使用できます。

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a>新規作成モード

この`recurrence`プロパティは、予定の繰り返しを管理するためのメソッドを提供する[定期的](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)なオブジェクトを返します。 これは予定に対して使用できます。

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var recurrence = asyncResult.value;
  if (!recurrence) {
    console.log("One-time appointment or meeting");
  } else {
    console.log(JSON.stringify(recurrence));
  }
}

// The following example shows the results of the getAsync call that retrieves the recurrence for a series.
// NOTE: In this example, seriesTimeObject is a placeholder for the JSON representing the
// recurrence.seriesTime property. You should use the SeriesTime object's methods to get the
// recurrence date and time properties.
Recurrence = {
  "recurrenceType": "weekly",
  "recurrenceProperties": {"interval": 2, "days": ["mon","thu","fri"], "firstDayOfWeek": "sun"},
  "seriesTime": {seriesTimeObject},
  "recurrenceTimeZone": {"name": "Pacific Standard Time", "offset": -480}
}
```

##### <a name="type"></a>型

* [繰り返さ](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)

|必要条件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.7|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|新規作成または閲覧|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a>requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)

イベントの必須出席者へのアクセスを提供します。 オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。

##### <a name="read-mode"></a>閲覧モード

`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。 既定では、コレクションは最大 100 人のメンバーに制限されています。 ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a>新規作成モード

`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。 既定では、コレクションは最大 100 人のメンバーに制限されています。 ただし、Windows および Mac では、次の制限が適用されます。

- 最大 500 人のメンバーを取得します。
- 呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a>型

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)

##### <a name="requirements"></a>Requirements

|必要条件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|新規作成または閲覧|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18"></a>sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)

電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。

メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsfrom) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。

> [!NOTE]
> `sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。

##### <a name="type"></a>型

*   [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)

##### <a name="requirements"></a>要件

|必要条件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|読み取り|

##### <a name="example"></a>例

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a>(nullable) 系列 Id: String

インスタンスが属する系列の id を取得します。

Web 上の Outlook およびデスクトップクライアントでは、 `seriesId`は、このアイテムが属する親 (シリーズ) アイテムの Exchange web サービス (EWS) ID を返します。 ただし、iOS と Android では、 `seriesId`は親アイテムの REST ID を返します。

> [!NOTE]
> `seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。 `seriesId`プロパティが OUTLOOK REST API で使用される outlook id と同じではありません。 この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。 詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api)」を参照してください。

この`seriesId`プロパティは`null` 、単一の予定、系列のアイテム、会議出席依頼などの親アイテムを持たないアイテムに`undefined`対して、会議出席依頼以外のアイテムに対して返されます。

##### <a name="type"></a>Type

* 文字列

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.7|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|新規作成または閲覧|

##### <a name="example"></a>例

```js
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-18"></a>start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)

予定を開始する日時を取得または設定します。

`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。

##### <a name="read-mode"></a>閲覧モード

`start` プロパティは `Date` オブジェクトを返します。

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a>新規作成モード

`start` プロパティは `Time` オブジェクトを返します。

[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。

次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。

```js
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

*   Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|新規作成または閲覧|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-18"></a>subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)

アイテムの件名フィールドに示される説明を取得または設定します。

`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。

##### <a name="read-mode"></a>閲覧モード

`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。

次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a>新規作成モード
`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a>型

*   String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|新規作成または閲覧|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a>to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)

メッセージの **To** 行にある受信者へのアクセスを提供します。 オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。

##### <a name="read-mode"></a>閲覧モード

`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。 既定では、コレクションは最大 100 人のメンバーに制限されています。 ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a>新規作成モード

`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。 既定では、コレクションは最大 100 人のメンバーに制限されています。 ただし、Windows および Mac では、次の制限が適用されます。

- 最大 500 人のメンバーを取得します。
- 呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a>型

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|新規作成または閲覧|

### <a name="methods"></a>メソッド

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a>addFileAttachmentAsync(uri, attachmentName, [options], [callback])

ファイルを添付ファイルとしてメッセージまたは予定に追加します。

`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。

その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。

##### <a name="parameters"></a>パラメーター
|名前|種類|属性|説明|
|---|---|---|---|
|`uri`|String||メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。|
|`attachmentName`|String||添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。|
|`options`|オブジェクト|&lt;オプション&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`|オブジェクト|&lt;optional&gt;|開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。|
|`options.isInline`|Boolean|&lt;省略可能&gt;|`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。|
|`callback`|function|&lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。 <br/>成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。<br/>添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。|

##### <a name="errors"></a>エラー

|エラー コード|説明|
|------------|-------------|
|`AttachmentSizeExceeded`|添付ファイルのサイズが上限を超えています。|
|`FileTypeNotSupported`|許可されていない拡張子の添付ファイルです。|
|`NumberOfAttachmentsExceeded`|メッセージまたは予定の添付ファイルが多すぎます。|

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.1|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|作成|

##### <a name="examples"></a>例

```js
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

次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。

```js
Office.context.mailbox.item.addFileAttachmentAsync(
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
        // Do something here.
      });
  });
```

<br>

---
---

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [callback])

Base64 エンコードのファイルを添付ファイルとしてメッセージまたは予定に追加します。

この`addFileAttachmentFromBase64Async`メソッドは、base64 エンコードからファイルをアップロードし、新規作成フォームのアイテムに添付します。 このメソッドは、AsyncResult オブジェクトの添付ファイル識別子を返します。

その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。

##### <a name="parameters"></a>パラメーター

|名前|型|属性|説明|
|---|---|---|---|
|`base64File`|String||電子メールまたはイベントに追加する画像またはファイルの、base64 でエンコードされたコンテンツ。|
|`attachmentName`|String||添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。|
|`options`|オブジェクト|&lt;オプション&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`|オブジェクト|&lt;オプション&gt;|開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。|
|`options.isInline`|Boolean|&lt;省略可能&gt;|`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。|
|`callback`|function|&lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。 <br/>成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。<br/>添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。|

##### <a name="errors"></a>エラー

|エラー コード|説明|
|------------|-------------|
|`AttachmentSizeExceeded`|添付ファイルのサイズが上限を超えています。|
|`FileTypeNotSupported`|許可されていない拡張子の添付ファイルです。|
|`NumberOfAttachmentsExceeded`|メッセージまたは予定の添付ファイルが多すぎます。|

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.8|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|作成|

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
        // Do something here.
      });
  });
```

<br>

---
---

#### <a name="addhandlerasynceventtype-handler-options-callback"></a>addHandlerAsync(eventType, handler, [options], [callback])

サポートされているイベントのイベント ハンドラーを追加します。

現在、サポートされて`Office.EventType.AttachmentsChanged`いる`Office.EventType.AppointmentTimeChanged`イベント`Office.EventType.EnhancedLocationsChanged`の`Office.EventType.RecipientsChanged`種類は`Office.EventType.RecurrenceChanged`、、、、、です。

##### <a name="parameters"></a>パラメーター

| 名前 | 型 | 属性 | 説明 |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || ハンドラーを呼び出す必要のあるイベント。 |
| `handler` | Function || イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。 |
| `options` | Object | &lt;オプション&gt; | 次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。 |
| `options.asyncContext` | オブジェクト | &lt;オプション&gt; | 開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。 |
| `callback` | function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.7 |
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧 |

##### <a name="example"></a>例

```js
function myHandlerFunction(eventarg) {
  if (eventarg.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added) {
    var attachment = eventarg.attachmentDetails;
    console.log("Event Fired and Attachment Added!");
    getAttachmentContentAsync(attachment.id, options, callback);
  }
}

Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, myHandlerFunction, myCallback);
```

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a>addItemAttachmentAsync(itemId, attachmentName, [options], [callback])

メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。

`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。

その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。

Office アドインを Outlook on the web で実行している場合、編集中のアイテム以外のアイテムに `addItemAttachmentAsync` メソッドでアイテムを添付できます。ただし、これはサポートされていないため、お勧めできません。

##### <a name="parameters"></a>パラメーター

|名前|型|属性|説明|
|---|---|---|---|
|`itemId`|String||添付するアイテムの Exchange 識別子。最大長は 100 文字です。|
|`attachmentName`|String||添付するアイテムの件名。 最大の長さは、255 文字です。|
|`options`|Object|&lt;オプション&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`|オブジェクト|&lt;省略可能&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`|関数|&lt;任意&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。 <br/>成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。<br/>添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。|

##### <a name="errors"></a>エラー

|エラー コード|説明|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|メッセージまたは予定の添付ファイルが多すぎます。|

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.1|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|作成|

##### <a name="example"></a>例

次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。

```js
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

<br>

---
---

#### <a name="close"></a>close()

作成中の現在の項目を閉じます。

`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。

> [!NOTE]
> Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。

Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.3|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|制限あり|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|新規作成|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a>displayReplyAllForm(formData, [callback])

選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。

> [!NOTE]
> このメソッドは、Outlook on iOS または Android ではサポートされていません。

Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。

文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。

`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。

##### <a name="parameters"></a>パラメーター

|名前|型|属性|説明|
|---|---|---|---|
|`formData`|String &#124; Object||回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。<br/>**または**<br/>本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。|
|`formData.htmlBody`|String|&lt;省略可能&gt;|回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。
|`formData.attachments`|Array.&lt;Object&gt;|&lt;省略可能&gt;|ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。|
|`formData.attachments.type`|String||添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。|
|`formData.attachments.name`|String||添付ファイル名を含む文字列。最大の長さは 255 文字です。|
|`formData.attachments.url`|文字列||`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。|
|`formData.attachments.isInline`|ブール値||`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。|
|`formData.attachments.itemId`|String||`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。|
|`callback`|function|&lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。|

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|読み取り|

##### <a name="examples"></a>例

次のコードは `displayReplyAllForm` 関数に文字列を渡します。

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

空の本文を返信します。

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

本文だけを返信します。

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

本文とファイルの添付ファイルを返信します。

```js
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

```js
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

```js
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

<br>

---
---

#### <a name="displayreplyformformdata-callback"></a>displayReplyForm(formData, [callback])

選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。

> [!NOTE]
> このメソッドは、Outlook on iOS または Android ではサポートされていません。

Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。

文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。

`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。

##### <a name="parameters"></a>パラメーター

|名前|型|属性|説明|
|---|---|---|---|
|`formData`|String &#124; Object||回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。<br/>**または**<br/>本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。|
|`formData.htmlBody`|String|&lt;省略可能&gt;|回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。
|`formData.attachments`|Array.&lt;Object&gt;|&lt;省略可能&gt;|ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。|
|`formData.attachments.type`|String||添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。|
|`formData.attachments.name`|String||添付ファイル名を含む文字列。最大の長さは 255 文字です。|
|`formData.attachments.url`|文字列||`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。|
|`formData.attachments.isInline`|ブール値||`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。|
|`formData.attachments.itemId`|String||`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。|
|`callback`|function|&lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。|

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|読み取り|

##### <a name="examples"></a>例

次のコードは `displayReplyForm` 関数に文字列を渡します。

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

空の本文を返信します。

```js
Office.context.mailbox.item.displayReplyForm({});
```

本文だけを返信します。

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

本文とファイルの添付ファイルを返信します。

```js
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

```js
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

```js
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

<br>

---
---

#### <a name="getallinternetheadersasyncoptions-callback"></a>getAllInternetHeadersAsync ([オプション], [callback])

メッセージのすべてのインターネットヘッダーを文字列として取得します。 閲覧モードのみ。

##### <a name="parameters"></a>パラメーター

|名前|型|属性|説明|
|---|---|---|---|
|`options`|オブジェクト|&lt;オプション&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`|オブジェクト|&lt;省略可能&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`|関数|&lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。 成功した場合、インターネットヘッダーデータは、文字列として asyncResult プロパティに提供されます。 返される文字列値の書式情報については、 [RFC 2183](https://tools.ietf.org/html/rfc2183)を参照してください。 呼び出しが失敗した場合、asyncResult. error プロパティには、エラーの理由と共にエラーコードが含まれます。|

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.8|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|読み取り|

##### <a name="returns"></a>戻り値:

[RFC 2183](https://tools.ietf.org/html/rfc2183)に従って書式設定された文字列としてのインターネットヘッダーデータ。

型:String

##### <a name="example"></a>例

```js
// Get the internet headers related to the mail.
Office.context.mailbox.item.getAllInternetHeadersAsync(
  function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log(asyncResult.value);
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is no context.
        // Treat as no context.
      } else {
        // Handle the error.
      }
    }
  }
);
```

<br>

---
---

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontentviewoutlook-js-18"></a>getAttachmentContentAsync (attachmentId, [options], [callback]) > [Attachmentcontent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)

メッセージまたは予定から指定された添付ファイルを取得し`AttachmentContent` 、それをオブジェクトとして返します。

メソッド`getAttachmentContentAsync`は、指定された id の添付ファイルをアイテムから取得します。 ベストプラクティスとして、識別子を使用して、または`getAttachmentsAsync` `item.attachments`の呼び出しで attachmentIds を取得したのと同じセッションの添付ファイルを取得する必要があります。 Outlook on the web とモバイル デバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。 ユーザーがアプリを閉じたとき、またはインラインフォームの作成が開始されたときに、別のウィンドウで続行するためにフォームをポップアウトした後、セッションが終了します。

##### <a name="parameters"></a>パラメーター

|名前|型|属性|説明|
|---|---|---|---|
|`attachmentId`|String||取得する添付ファイルの識別子を指定します。|
|`options`|オブジェクト|&lt;オプション&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`|オブジェクト|&lt;省略可能&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`|関数|&lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。|

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.8|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|新規作成または閲覧|

##### <a name="returns"></a>戻り値:

型: [Attachmentcontent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)

##### <a name="example"></a>例

```js
var item = Office.context.mailbox.item;
var listOfAttachments = [];
var options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      // Handle file attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Handle email item attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Handle .icalender attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      break;
    default:
      // Handle attachment formats that are not supported.
  }
}
```

<br>

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-18"></a>getAttachmentsAsync ([オプション], [callback]) > Array. <[Attachmentdetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)>

アイテムの添付ファイルを配列として取得します。 新規作成モードのみです。

##### <a name="parameters"></a>パラメーター

|名前|型|属性|説明|
|---|---|---|---|
|`options`|オブジェクト|&lt;オプション&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`|オブジェクト|&lt;省略可能&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`|function|&lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。|

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.8|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|作成|

##### <a name="returns"></a>戻り値:

型: Array. <[attachmentdetails 詳細](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)>

##### <a name="example"></a>例

次の例では、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。

```js
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      var attachment = result.value [i];
      outputString += "<BR>" + i + ". Name: ";
      outputString += attachment.name;
      outputString += "<BR>ID: " + attachment.id;
      outputString += "<BR>contentType: " + attachment.contentType;
      outputString += "<BR>size: " + attachment.size;
      outputString += "<BR>attachmentType: " + attachment.attachmentType;
      outputString += "<BR>isInline: " + attachment.isInline;
    }
  }
}
```

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-18"></a>getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}

選択したアイテムの本文にあるエンティティを取得します。

> [!NOTE]
> このメソッドは、Outlook on iOS または Android ではサポートされていません。

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|読み取り|

##### <a name="returns"></a>戻り値:

型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)

##### <a name="example"></a>例

次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-18meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-18phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-18tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-18"></a>getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}

選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。

> [!NOTE]
> このメソッドは、Outlook on iOS または Android ではサポートされていません。

##### <a name="parameters"></a>パラメーター

|名前|種類|説明|
|---|---|---|
|`entityType`|[Office.MailboxEnums.EntityType](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.8)|EntityType 列挙値の 1 つ。|

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|制限あり|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|読み取り|

##### <a name="returns"></a>戻り値:

`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。 指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。 それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。

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

型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>

##### <a name="example"></a>例

次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。

```js
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
};
```

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-18meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-18phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-18tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-18"></a>getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}

マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。

> [!NOTE]
> このメソッドは、Outlook on iOS または Android ではサポートされていません。

`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。

##### <a name="parameters"></a>パラメーター

|名前|種類|説明|
|---|---|---|
|`name`|String|一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。|

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|読み取り|

##### <a name="returns"></a>戻り値:

`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。

型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>

<br>

---
---

#### <a name="getitemidasyncoptions-callback"></a>getItemIdAsync ([オプション], callback)

保存されたアイテムの ID を非同期に取得します。 新規作成モードのみです。

このメソッドを呼び出すと、コールバックメソッドによってアイテム ID が返されます。

> [!NOTE]
> アドインが新規作成モードの`getItemIdAsync`アイテムに対して呼び出しを行う場合 ( `itemId` EWS または REST API を使用するため)、Outlook がキャッシュモードの場合は、アイテムがサーバーに同期されるまでしばらく時間がかかる場合があることに注意してください。 アイテムが同期されるまで、 `itemId`は認識されず、を使用するとエラーが返されます。

##### <a name="parameters"></a>パラメーター

|名前|型|属性|説明|
|---|---|---|---|
|`options`|オブジェクト|&lt;オプション&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`|オブジェクト|&lt;optional&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`|function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。<br/><br/>成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。|

##### <a name="errors"></a>エラー

|エラー コード|説明|
|------------|-------------|
|`ItemNotSaved`|この id は、アイテムが保存されるまでは取得できません。|

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.8|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|作成|

##### <a name="examples"></a>例

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

次の例は、コールバック関数`result`に渡されるパラメーターの構造を示しています。 プロパティ`value`には、アイテムの ID が含まれています。

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a>getRegExMatches() → {Object}

選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。

> [!NOTE]
> このメソッドは、Outlook on iOS または Android ではサポートされていません。

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

アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|読み取り|

##### <a name="returns"></a>戻り値:

マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。

<dl class="param-type">

<dt>型</dt>

<dd>Object</dd>

</dl>

##### <a name="example"></a>例

次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a>getRegExMatchesByName(name) → (nullable) {Array.< String >}

選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。

> [!NOTE]
> このメソッドは、Outlook on iOS または Android ではサポートされていません。

`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。

アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。

##### <a name="parameters"></a>パラメーター

|名前|種類|説明|
|---|---|---|
|`name`|String|一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。|

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|読み取り|

##### <a name="returns"></a>戻り値:

マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。

型: Array.< String >

##### <a name="example"></a>例

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a>getSelectedDataAsync(coercionType, [options], callback) → {String}

メッセージの件名または本文から非同期的に選択したデータを返します。

選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。

> [!NOTE]
> Web 上の Outlook では、テキストが選択されておらず、カーソルが本文にある場合、このメソッドは文字列 "null" を返します。 このような状況を確認するには、次のようなコードを含めます。
>
> `var selectedText = (asyncResult.value.endPosition === asyncResult.value.startPosition) ? "" : asyncResult.value.data;`

##### <a name="parameters"></a>パラメーター

|名前|型|属性|説明|
|---|---|---|---|
|`coercionType`|[Office.CoercionType](office.md#coerciontype-string)||データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。|
|`options`|オブジェクト|&lt;オプション&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`|オブジェクト|&lt;optional&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`|function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。<br/><br/>コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。 選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。|

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.2|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|作成|

##### <a name="returns"></a>戻り値:

選択されたデータ (`coercionType` で決定された形式の文字列)。

型:String

##### <a name="example"></a>例

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
  // Check for errors.
}
```

<br>

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-18"></a>getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}

強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。 強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。

> [!NOTE]
> このメソッドは、Outlook on iOS または Android ではサポートされていません。

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.6|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|読み取り|

##### <a name="returns"></a>戻り値:

型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)

##### <a name="example"></a>例

次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a>getSelectedRegExMatches() → {Object}

マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。

> [!NOTE]
> このメソッドは、Outlook on iOS または Android ではサポートされていません。

`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。

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

アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.6|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|読み取り|

##### <a name="returns"></a>戻り値:

マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。

##### <a name="example"></a>例

次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a>getSharedPropertiesAsync ([options], callback)

共有フォルダー、予定表、またはメールボックス内の選択した予定またはメッセージのプロパティを取得します。

##### <a name="parameters"></a>パラメーター

|名前|型|属性|説明|
|---|---|---|---|
|`options`|オブジェクト|&lt;オプション&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`|オブジェクト|&lt;optional&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`|function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。<br/><br/>共有プロパティは、 [`SharedProperties`](/javascript/api/outlook/office.sharedproperties?view=outlook-js-1.8) `asyncResult.value`プロパティのオブジェクトとして提供されます。 このオブジェクトは、アイテムの共有プロパティを取得するために使用できます。|

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.8|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|新規作成または閲覧|

##### <a name="example"></a>例

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a>loadCustomPropertiesAsync(callback, [userContext])

選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。

カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。

##### <a name="parameters"></a>パラメーター

|名前|型|属性|説明|
|---|---|---|---|
|`callback`|function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。<br/><br/>カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.8) オブジェクトとして指定されます。 このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。|
|`userContext`|オブジェクト|&lt;省略可能&gt;|開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。 このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。|

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|新規作成または閲覧|

##### <a name="example"></a>例

次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。

```js
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(customPropsCallback);
  });
};

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

<br>

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a>removeAttachmentAsync(attachmentId, [options], [callback])

メッセージまたは予定から添付ファイルを削除します。

`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。 ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。 Outlook on the web とモバイル デバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。 ユーザーがアプリを閉じたとき、またはインラインフォームの作成が開始されたときに、別のウィンドウで続行するためにフォームをポップアウトした後、セッションが終了します。

##### <a name="parameters"></a>パラメーター

|名前|型|属性|説明|
|---|---|---|---|
|`attachmentId`|String||削除する添付ファイルの識別子。|
|`options`|オブジェクト|&lt;オプション&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`|オブジェクト|&lt;省略可能&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`|function|&lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。 <br/>添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。|

##### <a name="errors"></a>エラー

|エラー コード|説明|
|------------|-------------|
|`InvalidAttachmentId`|添付ファイル識別子が存在しません。|

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.1|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|作成|

##### <a name="example"></a>例

次のコードは、'0' の識別子を持つ添付ファイルを削除します。

```js
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

<br>

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a>removeHandlerAsync(eventType, [options], [callback])

サポートされているイベントの種類のイベント ハンドラーを削除します。

現在、サポートされて`Office.EventType.AttachmentsChanged`いる`Office.EventType.AppointmentTimeChanged`イベント`Office.EventType.EnhancedLocationsChanged`の`Office.EventType.RecipientsChanged`種類は`Office.EventType.RecurrenceChanged`、、、、、です。

##### <a name="parameters"></a>パラメーター

| 名前 | 型 | 属性 | 説明 |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || ハンドラーを取り消すイベント。 |
| `options` | オブジェクト | &lt;オプション&gt; | 次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。 |
| `options.asyncContext` | オブジェクト | &lt;省略可能&gt; | 開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。 |
| `callback` | function| &lt;optional&gt;|メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.7 |
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧 |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a>saveAsync([options], callback)

項目を非同期的に保存します。

呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。 Outlook on the web またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。 キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。

> [!NOTE]
> EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。 アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。

予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。

> [!NOTE]
> 次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。
>
> - Outlook on Mac では、会議の保存はサポートされていません。 `saveAsync` メソッドは、作成モードの会議から呼び出されると失敗します。 回避策については、「[Office JS API を使用して Outlook for Mac で会議を下書きとして保存できない](https://support.microsoft.com/help/4505745)」を参照してください。
> - Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。

##### <a name="parameters"></a>パラメーター

|名前|型|属性|説明|
|---|---|---|---|
|`options`|オブジェクト|&lt;オプション&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`|オブジェクト|&lt;省略可能&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`|function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。<br/><br/>成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。|

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.3|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|作成|

##### <a name="examples"></a>例

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a>setSelectedDataAsync(data, [options], callback)

メッセージの本文または件名に非同期的にデータを挿入します。

`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。

##### <a name="parameters"></a>パラメーター

|名前|型|属性|説明|
|---|---|---|---|
|`data`|String||挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。|
|`options`|Object|&lt;オプション&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`|オブジェクト|&lt;任意&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`options.coercionType`|[Office.CoercionType](office.md#coerciontype-string)|&lt;optional&gt;|`text` の場合、Outlook on the web とデスクトップ クライアントでは現在のスタイルが適用されます。 フィールドが HTML エディターの場合、データが HTML の場合でもテキスト データのみが挿入されます。<br/><br/>`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Outlook on the web では現在のスタイルが適用され、Outlook デスクトップ クライアントでは既定のスタイルが適用されます。 フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。<br/><br/>`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。|
|`callback`|function||メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。|

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.2|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)|作成|

##### <a name="example"></a>例

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```