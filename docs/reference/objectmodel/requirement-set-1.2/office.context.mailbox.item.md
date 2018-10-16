
# <a name="item"></a>項目

### [Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item` 名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予約にアクセスします。 `item`の種類を[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype)プロパティを使用して決定できます。

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| 制限あり|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

### <a name="example"></a>例

次の JavaScript のコード例は、Outlook の現在の項目の `subject` プロパティにアクセスする方法を示しています。

```JavaScript
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

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a>添付ファイル :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)>

項目の添付ファイルの配列を取得します。閲覧モード専用です。

> [!NOTE]
> 潜在的なセキュリティ問題により、特定の種類のファイルは Outlook でブロックされ、したがって戻ってきません。 詳細については、[「Outlook でブロックされた添付ファイル」](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)をご覧下さい。

##### <a name="type"></a>種類:

*   配列.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)>

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

次のコードでは、現在の項目にあるすべての添付ファイルの詳細を含む HTML 文字列を作成します。

```JavaScript
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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a>BCC:[受信者](/javascript/api/outlook_1_2/office.recipients)

メッセージの BCC (ブラインド カーボン コピー) 列上の 受信者を取得または更新するメソッドを提供するオブジェクトを取得します。 作成モード専用。

##### <a name="type"></a>種類:

*   [受信者](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成|

##### <a name="example"></a>例

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a>本文:[本文](/javascript/api/outlook_1_2/office.body)

項目の本文を操作するメソッドを提供するオブジェクトを取得します。

##### <a name="type"></a>種類:

*   [本文](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a>CC:配列.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_2/office.recipients)

メッセージの CC (カーボン コピー) 受信者へのアクセスを提供します。 オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。

##### <a name="read-mode"></a>閲覧モード

|||UNTRANSLATED_CONTENT_START|||The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.|||UNTRANSLATED_CONTENT_END|||

##### <a name="compose-mode"></a>作成モード

|||UNTRANSLATED_CONTENT_START|||The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.|||UNTRANSLATED_CONTENT_END|||

##### <a name="type"></a>種類:

*   配列.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>| [受信者](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="example"></a>例

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a>（Null 許容）conversationId:文字列

特定のメッセージが含まれている電子メールの会話の識別子を取得します。

メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティの整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。

新規作成フォームで新しい項目に対してこのプロパティに null を取得します。ユーザーが件名を設定しアイテムを保存する場合、`conversationId` プロパティは値を返します。

##### <a name="type"></a>種類:

*   文字列

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

#### <a name="datetimecreated-date"></a>dateTimeCreated:日付

アイテムが作成された日時を取得します。閲覧モード専用です。

##### <a name="type"></a>種類:

*   日付

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a>dateTimeModified:日付

項目が最後に変更された日時を取得します。閲覧モード専用です。

> [!NOTE]
> このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。

##### <a name="type"></a>種類:

*   日付

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a>end:日付|[時間](/javascript/api/outlook_1_2/office.time)

予約が終了する日時を取得または設定します。

`end`プロパティは、協定世界時 (UTC) 形式の時刻値として表示されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。

##### <a name="read-mode"></a>閲覧モード

`end`プロパティは`Date`オブジェクトを返します。

##### <a name="compose-mode"></a>作成モード

`end`プロパティは`Time`オブジェクトを返します。

[`Time.setAsync` ](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアントが所在するローカルの時刻をサーバーの UTC に変換する必要があります。

##### <a name="type"></a>種類:

*   日付 | [時間](/javascript/api/outlook_1_2/office.time)

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="example"></a>例

次の例では、[ オブジェクトの`setAsync`  ](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) `Time`  メソッドを使用して、作成モードで予定の終了時刻を設定します。

```JavaScript
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a>from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)

メッセージの送信者の電子メール アドレスを取得します。閲覧モード専用です。

メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) プロパティは同一人物を表します。その場合、`from` プロパティは委任者を、送信者プロパティは代理人を表します。

> [!NOTE]
> `recipientType`  プロパティ内の`EmailAddressDetails`    オブジェクトの`from`   プロパティは、 `undefined`です。

##### <a name="type"></a>種類:

*   [EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

#### <a name="internetmessageid-string"></a>internetMessageId: 文字列

電子メール メッセージ用のインターネット メッセージの識別子を取得します。閲覧モード専用です。

##### <a name="type"></a>種類:

*   文字列

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a>itemClass:文字列

選択された項目の Exchange Web サービスの項目 クラスを取得します。閲覧モード専用です。

`itemClass`プロパティは、選択したアイテムのメッセージ クラスを指定します。次は、メッセージまたは予定アイテムの既定のメッセージ クラスを示しています。

| 種類 | 説明 | 項目 クラス |
| --- | --- | --- |
| 予定項目 | これらは、項目クラス `IPM.Appointment`または`IPM.Appointment.Occurence`の予定表アイテムです。 | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| メッセージの項目
 | これには、基本のメッセージ クラス として`IPM.Note`を使用する、既定のメッセージ クラス`IPM.Schedule.Meeting`会議出席依頼、返信および取り消しを含む電子メール メッセージが含まれます。 | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

既定のメッセージ クラスを拡張したカスタム メッセージ クラス 、たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` などを作成することができます。

##### <a name="type"></a>種類:

*   文字列

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a>(Null 許容)itemId: 文字列

現在の項目の Exchange Web サービスのアイテム識別子を取得します。閲覧モード専用です。

> [!NOTE]
> `itemId` プロパティから返される識別子は、Exchange Web サービスの項目識別子と同じです。 `itemId` プロパティは、Outlook Entry ID または Outlook REST API によって使用される ID と同じではありません。 この値を使用して REST API の呼び出しを行う前に、 必要な設定1.3から利用可能な `Office.context.mailbox.convertToRestId`を使用して変換する必要があります。 詳細については、[「Outlook アドインから Outlook REST API の使用」](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)をご覧下さい。

##### <a name="type"></a>種類:

*   文字列

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

次のコードは項目識別子のプレゼンスを確認します。`itemId` プロパティが `null` または `undefined` を返す場合、項目はストアに保存され、非同期の結果から項目識別子が取得されます。

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a>itemType:[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

インスタンスが表している項目の種類を取得します。

`itemType`プロパティは、`ItemType`列挙値の 1 つを返します。これは`item`オブジェクト インスタンスがメッセージまたは予定のどちらであるかを示すものです。

##### <a name="type"></a>種類:

*   [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="example"></a>例

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a>場所: 文字列 | [場所](/javascript/api/outlook_1_2/office.location)

予約の場所を取得または設定します。

##### <a name="read-mode"></a>閲覧モード

`location`プロパティは、予定の場所を含む文字列を返します。

##### <a name="compose-mode"></a>作成モード

`location`プロパティは、予定の場所を取得または設定するために使用するメソッドを提供する`Location`オブジェクトを返します。

##### <a name="type"></a>種類:

*   文字列 | [場所](/javascript/api/outlook_1_2/office.location)

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="example"></a>例

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a>normalizedSubject: 文字列

すべてのプレフィックス (`RE:` や `FWD:` を含む) が削除された項目の件名を取得します。閲覧モード専用です。

normalizedSubject プロパティは、電子メール プログラムから追加された標準のプレフィックス (`RE:` や `FW:` など) が付く項目の件名を取得します。プレフィックスが付いたままの状態で項目の件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) プロパティを使用します。

##### <a name="type"></a>種類:

*   文字列

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a>optionalAttendees: 配列.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_2/office.recipients)

オプションのイベント出席者へのアクセスを提供します。 オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。

##### <a name="read-mode"></a>閲覧モード

`optionalAttendees`プロパティは、会議への各任意出席者の`EmailAddressDetails`オブジェクトを含む配列を返します。

##### <a name="compose-mode"></a>作成モード

`optionalAttendees`プロパティは会議への任意出席者を取得または設定するためのメソッドを提供する`Recipients`オブジェクトを返します。

##### <a name="type"></a>種類:

*   配列.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>| [受信者](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="example"></a>例

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a>開催者:[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)

特定の会議開催者の電子メールアドレスを取得します。閲覧モード専用です。

##### <a name="type"></a>種類:

*   [EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a>requiredAttendees: 配列.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_2/office.recipients)

イベントの必須出席者へのアクセスを提供します。 オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。

##### <a name="read-mode"></a>閲覧モード

`requiredAttendees`プロパティは、会議への各必須出席者の`EmailAddressDetails`オブジェクトを含む配列を返します。

##### <a name="compose-mode"></a>作成モード

`requiredAttendees`プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する`Recipients`オブジェクトを返します。

##### <a name="type"></a>種類:

*   配列.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>| [受信者](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="example"></a>例

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a>送信者:[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)

電子メール送信者のメールアドレスを取得します。閲覧モード専用です。

メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails)プロパティと`sender`プロパティは同一人物を表します。その場合、`from`プロパティは委任者を、送信者プロパティは代理人を表します。

> [!NOTE]
> `recipientType`  プロパティ内の`EmailAddressDetails`    オブジェクトの`sender`   プロパティは、 `undefined`です。

##### <a name="type"></a>種類:

*   [EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a>開始: 日付 | [時間](/javascript/api/outlook_1_2/office.time)

予定を開始する日時を取得または設定します。

`start`プロパティは、協定世界時 (UTC) 形式の時刻値として表示されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime)メソッドを使用して、値をクライアントのローカル日時に変換することができます。

##### <a name="read-mode"></a>閲覧モード

`start`プロパティは`Date`オブジェクトを返します。

##### <a name="compose-mode"></a>作成モード

`start`プロパティは`Time`オブジェクトを返します。

[`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-)メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアントが所在するローカル時刻をサーバーの UTC に変換する必要があります。

##### <a name="type"></a>種類:

*   日付 | [時間](/javascript/api/outlook_1_2/office.time)

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="example"></a>例

次の例では、[ オブジェクトの`setAsync` ](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-)  `Time`  メソッドを使用して、新規作成モードで予定の開始時刻を設定します。

```JavaScript
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

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a>件名: 文字列 | [件名](/javascript/api/outlook_1_2/office.subject)

項目の件名フィールドに表示される説明を取得または設定します。

`subject`プロパティは、電子メールサーバーから送信された項目の全件名を取得または設定します。

##### <a name="read-mode"></a>閲覧モード

`subject`プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string)プロパティを使用して、行間にある`RE:`や`FW:`のなどのすべてのプレフィックスを削除した件名を取得します。

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a>作成モード

`subject`プロパティは、件名を取得または設定するためのメソッドを提供する`Subject`オブジェクトを返します。

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a>種類:

*   文字列 | [件名](/javascript/api/outlook_1_2/office.subject)

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a>宛先: 配列.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_2/office.recipients)

メッセージの **宛先**列にある受信者へのアクセスを提供します。 オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。

##### <a name="read-mode"></a>閲覧モード

`to`プロパティは、メッセージの `EmailAddressDetails` 宛先**  列一覧にある各受信者の **  オブジェクトを含む配列を返します。コレクションのメンバーは 100 個までに制限されています。

##### <a name="compose-mode"></a>作成モード

`to`プロパティは、メッセージの `Recipients` 宛先**  列にある受信者を取得または更新するメソッドを提供する** オブジェクトを返します。

##### <a name="type"></a>種類:

*   配列.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>| [受信者](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="example"></a>例

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a>メソッド

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a>addFileAttachmentAsync(uri, attachmentName, [オプション], [コールバック])

ファイルを添付ファイルとしてメッセージまたは予定に追加します。

`addFileAttachmentAsync`メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内の項目に添付します。

その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)メソッドで識別子を使用して同じセッションの添付ファイルを削除することができます。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 属性| 説明|
|---|---|---|---|
|`uri`| 文字列||メッセージまたは予定に添付するファイルの場所を示す URIです。 2048 文字以内で入力してください。|
|`attachmentName`| 文字列||添付ファイルのアップロード時に表示される添付ファイルの名前です。 255 文字以内で入力してください。|
|`options`| オブジェクト| &lt;オプション&gt;|以下のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`| オブジェクト| &lt;オプション&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`| 関数| &lt;オプション&gt;|メソッドが完了すると、`callback`パラメータで渡された関数が、シングル パラメータ、`asyncResult`、で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。 <br/>これが成功すると、添付ファイルの識別子が`asyncResult.value`プロパティに提供されます。<br/>添付ファイルのアップロードに失敗した場合、`asyncResult`オブジェクトには、エラーの説明を提供する`Error`オブジェクトが含まれます。|

##### <a name="errors"></a>エラー

| エラー コード | 説明 |
|------------|-------------|
| `AttachmentSizeExceeded` | 添付ファイルのサイズが大きすぎます。
 |
| `FileTypeNotSupported` | 許可されていない拡張子の添付ファイルです。 |
| `NumberOfAttachmentsExceeded` | メッセージまたは予定の添付ファイルが多すぎます。 |

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成|

##### <a name="example"></a>例

```JavaScript
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a>addItemAttachmentAsync (itemd、attachmentName、[オプション]、[コールバック])

メッセージなどの Exchange 項目を添付ファイルとして、メッセージまたは予定に追加します。

`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つ項目を新規作成フォーム内の項目に添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメータがあるメソッドが呼び出されます。このパラメータには、添付ファイルの識別子、または項目の添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメータを使用して、状態情報をコールバック メソッドに渡すことができます。

その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)メソッドで識別子を使用して同じセッションの添付ファイルを削除することができます。

Office アドインを Outlook Web アプリケーションで実行している場合、`addItemAttachmentAsync`メソッドは項目を、編集中以外の項目に添付できますが、これはサポートされておらず、推奨されていません。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 属性| 説明|
|---|---|---|---|
|`itemId`| 文字列||添付する項目の Exchange 識別子です。100 文字以内で入力してください。|
|`attachmentName`| 文字列||添付する項目の件名です。 255 文字以内で入力してください。|
|`options`| オブジェクト| &lt;オプション&gt;|以下のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`| オブジェクト| &lt;オプション&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`| 関数| &lt;オプション&gt;|メソッドが完了すると、`callback`パラメータで渡された関数が、シングル パラメータ、`asyncResult`、で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。 <br/>これが成功すると、添付ファイルの識別子が`asyncResult.value`プロパティに提供されます。<br/>添付ファイルの追加に失敗した場合、 エラーの説明を提供する`asyncResult`オブジェクトが`Error`オブジェクトに含まれます。|

##### <a name="errors"></a>エラー

| エラー コード | 説明 |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | メッセージまたは予定の添付ファイルが多すぎます。 |

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成|

##### <a name="example"></a>例

次の例では、既存の Outlook 項目を名前付き `My Attachment` の添付ファイルとして追加されます。

```JavaScript
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

#### <a name="displayreplyallformformdata"></a>displayReplyAllForm(formData)

選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した返信フォームが表示されます。

> [!NOTE]
> このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。

Outlook Web アプリでは、返信フォームは、3 列ビューのポップアウト形式、および 2 列または 1 列ビューのポップアップ形式で表示されます。

文字列パラメータのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。

`formData.attachments` パラメータで添付ファイルを指定すると、Outlook と Office Web アプリ はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 説明|
|---|---|---|
|`formData`| 文字列 | オブジェクト| |返信フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。<br/>**または**<br/>本文または添付ファイルのデータと、コールバック関数を含むオブジェクトです。オブジェクトの定義は次のとおりです。 |
| `formData.htmlBody` | 文字列 | &lt;オプション&gt; | 返信フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。
| `formData.attachments` | 配列。&lt;オブジェクト&gt; | &lt;オプション&gt; | ファイルまたは項目の添付ファイルである JSON オブジェクトの配列です。 |
| `formData.attachments.type` | 文字列 | | 添付ファイルの種類を示します。添付ファイルの場合は`file`、添付項目の場合は`item`でなければなりません。 |
| `formData.attachments.name` | 文字列 | | 添付ファイル名を含む文字列で、255 文字以内で入力が可能です。|
| `formData.attachments.url` | 文字列 | | `type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URIです。 |
| `formData.attachments.itemId` | 文字列 | | `type`が`item`に設定されている場合にのみ使用されます。添付ファイルの EWS アイテム ID です。 100 文字以内の文字列です。 |
| `callback` | 関数 | &lt;オプション&gt; | メソッドが完了すると、 `callback` パラメータに渡された関数が、シングル パラメータ, `asyncResult`で呼び出されます。これは [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトです。 |

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="examples"></a>例

次のコードは`displayReplyAllForm`関数に文字列を渡します。

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

空の本文を返信します。

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

本文だけを返信します。

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

本文と添付ファイルを返信します。

```JavaScript
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

本文と項目の添付ファイルを返信します。

```JavaScript
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

本文、添付ファイル、項目の添付ファイル、およびコールバックを返信します。

```JavaScript
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

選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む返信フォームが表示されます。

> [!NOTE]
> このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。

Outlook Web アプリでは、返信フォームは、3 列ビューのポップアウト形式、および 2 列または 1 列ビューのポップアップ形式で表示されます。

文字列パラメータのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。

`formData.attachments` パラメータで添付ファイルを指定すると、Outlook と Office Web アプリ はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 説明|
|---|---|---|
|`formData`| 文字列 | オブジェクト| | 返信フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。<br/>**または**<br/>本文または添付ファイルのデータと、コールバック関数を含むオブジェクトです。オブジェクトの定義は次のとおりです。 |
| `formData.htmlBody` | 文字列 | &lt;オプション&gt; | 返信フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。
| `formData.attachments` | 配列。&lt;オブジェクト&gt; | &lt;オプション&gt; | ファイルまたは項目の添付ファイルである JSON オブジェクトの配列です。 |
| `formData.attachments.type` | 文字列 | | 添付ファイルの種類を示します。添付ファイルの場合は`file`、添付項目の場合は`item`でなければなりません。 |
| `formData.attachments.name` | 文字列 | | 添付ファイル名を含む文字列で、255 文字以内で入力が可能です。|
| `formData.attachments.url` | 文字列 | | `type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URIです。 |
| `formData.attachments.itemId` | 文字列 | | `type`が`item`に設定されている場合にのみ使用されます。添付ファイルの EWS アイテム ID です。 100 文字以内の文字列です。 |
| `callback` | 関数 | &lt;オプション&gt; | メソッドが完了すると、 `callback` パラメータに渡された関数が、シングル パラメータ, `asyncResult`で呼び出されます。これは [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトです。 |

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="examples"></a>例

次のコードは`displayReplyForm`関数に文字列を渡します。

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

空の本文を返信します。

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

本文だけを返信します。

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

本文と添付ファイルを返信します。

```JavaScript
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

本文と項目の添付ファイルを返信します。

```JavaScript
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

本文、添付ファイル、項目の添付ファイル、およびコールバックを返信します。

```JavaScript
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

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a>getEntities() → {[エンティティ](/javascript/api/outlook_1_2/office.entities)}

選択した項目の本文で見つかったエンティティを取得します。

> [!NOTE]
> このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="returns"></a>戻り値:

種類: [エンティティ](/javascript/api/outlook_1_2/office.entities)

##### <a name="example"></a>例

次の例では、現在の項目の本文内にある連絡先のエンティティにアクセスします。

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a>getEntitiesByType(entityType)] → [(Null 許容) {配列<(文字列 |[連絡先](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion)) >}

選択した項目で見つかった指定のエンティティ型のエンティティすべてを含む配列を取得します。

> [!NOTE]
> このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 説明|
|---|---|---|
|`entityType`| [Office.MailboxEnums.EntityType](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|EntityType 列挙値の 1 つです。|

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| 制限あり|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="returns"></a>戻り値:

 `entityType` に渡された値が有効な `EntityType` 列挙型のメンバーでない場合、メソッドは Nullを返します。 指定した種類のエンティティが項目の本文に存在しない場合、メソッドは空の配列を返します。 そうでない場合、返される配列内のオブジェクトの種類は、 `entityType` パラメータ内で要求されたエンティティの種類によって異なります。

このメソッドを使用する最小限のアクセス許可のレベルは **制限あり**ですが、一部のエンティティには、次のテーブルで指定されているように、アクセスに **ReadItem** が必要です。

| の値 `entityType` | 返される配列内にあるオブジェクトの種類 | 必要なアクセス許可のレベル |
| --- | --- | --- |
| `Address` | 文字列 | **制限あり** |
| `Contact` | 連絡先 | **ReadItem** |
| `EmailAddress` | 文字列 | **ReadItem** |
| `MeetingSuggestion` | MeetingSuggestion | **ReadItem** |
| `PhoneNumber` | PhoneNumber | **制限あり** |
| `TaskSuggestion` | TaskSuggestion | **ReadItem** |
| `URL` | 文字列 | **制限あり** |

種類: 配列.<(文字列 |[連絡先](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>

##### <a name="example"></a>例

次の例は、現在の項目の本文にある郵便アドレスを表す文字列の配列にアクセスする方法を示します。

```JavaScript
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a>getFilteredEntitiesByName(name)] → [(Null 許容) {配列<(文字列| [連絡先](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[電話番号 ](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}

 選択済み項目の既知のエンティティを返し、 XML ファイルで定義された名前付きフィルターを渡します。

> [!NOTE]
> このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。

`getFilteredEntitiesByName` メソッドは、指定された [    要素値があるマニフェストXMLファイル内のルール要素 ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule)  で定義された正規表現に一致するエンティティを返します。`FilterName`

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 説明|
|---|---|---|
|`name`| 文字列|一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前です。|

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="returns"></a>戻り値:

 `ItemHasKnownEntity`  パラメータと一致する `FilterName`  要素値を持つ `name`  要素がマニフェスト内にない場合、メソッドは `null`を返します。  `name` パラメータがマニフェスト内の `ItemHasKnownEntity` 要素と一致するが、現在の項目内に一致するエンティティがない場合は、メソッドは空の配列を返します。

種類: 配列.<(文字列 |[連絡先](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>

#### <a name="getregexmatches--object"></a>getRegExMatches() → {オブジェクト}

マニフェスト XML ファイルで定義された正規表現に一致する選択済みの項目の文字列の値を返します。

> [!NOTE]
> このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。

 `getRegExMatches` メソッドは、マニフェスト XML ファイル内にある各 `ItemHasRegularExpressionMatch` または`ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定された項目のプロパティに一致しなければなりません。 `PropertyName` 単純型は、サポートされるプロパティを定義します。

たとえば、アドイン マニフェストに次の `Rule` 要素があると見なします。

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

`getRegExMatches` から返されるオブジェクトには、 `fruits` および `veggies`という 2 つのプロパティがあります。

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> 項目の本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、項目の本文全体を返さないようにします。`.*` などの正規表現を使用して項目の本文全体を取得しても、期待する結果が返されないことがあります。

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="returns"></a>戻り値:

マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列を含むオブジェクトです。各配列の名前は、一致する `RegExName`  ルールの `ItemHasRegularExpressionMatch`  属性または一致する `FilterName`   ルールの `ItemHasKnownEntity`  属性の対応する値と等しくなります。

<dl class="param-type">

<dt>種類</dt>

<dd>オブジェクト</dd>

</dl>

##### <a name="example"></a>例

次の例は、マニフェストで指定された正規表現の <rule> 要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a>getRegExMatchesByName(name)] → [(Null許容) {配列. < 文字列 >}

選択した項目内の文字列を返し、マニフェスト XML ファイルで定義された名前付きの正規表現に一致します。

> [!NOTE]
> このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。

`getRegExMatchesByName` メソッドは、指定された`ItemHasRegularExpressionMatch`  要素値を持つマニフェスト XML ファイルの`RegExName`  ルール要素で定義された正規表現に一致する文字列を返します。

項目の本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、項目の本文全体を返さないようにします。`.*` などの正規表現を使用して項目の本文全体を取得しても、期待する結果が返されないことがあります。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 説明|
|---|---|---|
|`name`| 文字列|一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前です。|

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="returns"></a>戻り値:

マニフェスト XML ファイルで定義された正規表現に一致する文字列を含む配列です。

<dl class="param-type">

<dt>種類</dt>

<dd>配列. < 文字列 ></dd>

</dl>

##### <a name="example"></a>例

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a>getSelectedDataAsync(coercionType, [オプション], コールバック) → {文字列}

メッセージの件名または本文から非同期的に選択したデータを返します。

選択したデータがなく、カーソルが本文または件名にある場合、メソッドは選択したデータに対して Null を返します。本文または件名以外のフィールドが選択されている場合、メソッドは `InvalidSelection` エラーを返します。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 属性| 説明|
|---|---|---|---|
|`coercionType`| [Office.CoercionType](office.md#coerciontype-string)||データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーン テキストまたは HTML のいずれの場合も選択されたテキストを返します。|
|`options`| オブジェクト| &lt;オプション&gt;|以下のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`| オブジェクト| &lt;オプション&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`| 関数||メソッドが完了すると、`callback`パラメータで渡された関数が、シングル パラメータ、`asyncResult`、で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。<br/><br/>コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data`を呼び出します。 選択元のソース プロパティにアクセスするには、 `asyncResult.value.sourceProperty` を呼び出します。これは `body`または `subject`になります。|

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.2以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成|

##### <a name="returns"></a>戻り値:

`coercionType`で決定された書式設定の文字列として選択されたデータです。

<dl class="param-type">

<dt>種類</dt>

<dd>文字列</dd>

</dl>

##### <a name="example"></a>例

```JavaScript
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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a>loadCustomPropertiesAsync(callback, [userContext])

選択された項目のアドインのカスタム プロパティを非同期に読み込みます。

カスタム プロパティは、アプリごと、項目ごとにキーと値のペアとして保管されます。このメソッドは、コールバックで  `CustomProperties` オブジェクトを返します。このオブジェクトは、現在の項目および現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、項目上では暗号化されないので、安全な保管場所として使用しないでください。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 属性| 説明|
|---|---|---|---|
|`callback`| 関数||メソッドが完了すると、`callback`パラメータで渡された関数が、シングル パラメータ、`asyncResult`、で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。<br/><br/>カスタム プロパティは [  プロパティの `CustomProperties`  ](/javascript/api/outlook_1_2/office.customproperties) `asyncResult.value`  オブジェクトとして指定されます。 このオブジェクトは、項目からカスタム プロパティを取得、設定、および削除し、カスタム プロパティに対する変更をサーバーに設定し直すために使用することができます。|
|`userContext`| オブジェクト| &lt;オプション&gt;|開発者は、コールバック関数でアクセスしたいオブジェクトを提供することができます。 このオブジェクトは、コールバック関数の`asyncResult.asyncContext`プロパティによってアクセスすることができます。|

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="example"></a>例

次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、 `CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティを読み込んだ後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp`を読み取り、 `CustomProperties.set` メソッドでカスタム プロパティ `otherProp`を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。

```JavaScript
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a>removeAttachmentAsync(attachmentId、[オプション]、 [コールバック])

メッセージまたは予定から添付ファイルを削除します。

`removeAttachmentAsync`メソッドは、指定した識別子の添付ファイルを項目 から削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web アプリ とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別のウィンドウで操作を継続すると、セッションは終了します。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 属性| 説明|
|---|---|---|---|
|`attachmentId`| 文字列||削除する添付ファイルの識別子です。配列は 100 文字以内で入力してください。|
|`options`| オブジェクト| &lt;オプション&gt;|以下のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`| オブジェクト| &lt;オプション&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`| 関数| &lt;オプション&gt;|メソッドが完了すると、`callback`パラメータで渡された関数が、シングル パラメータ、`asyncResult`、で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。 <br/>添付ファイルの削除に失敗すると、`asyncResult.error`プロパティにはエラー コードとエラーの理由が含まれます。|

##### <a name="errors"></a>エラー

| エラー コード | 説明 |
|------------|-------------|
| `InvalidAttachmentId` | 添付ファイルの識別子が存在しません。 |

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.1以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成|

##### <a name="example"></a>例

次のコードは、「0」の識別子を持つ添付ファイルを削除します。

```JavaScript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="setselecteddataasyncdata-options-callback"></a>setSelectedDataAsync(日付、 [オプション]、 コールバック)

メッセージの本文または件名に非同期的にデータを挿入します。

`setSelectedDataAsync`メソッドは、指定された文字列を項目のサブジェクトまたは本文のカーソル位置に挿入するか、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 属性| 説明|
|---|---|---|---|
|`data`| 文字列||挿入されるデータです。データの長さは 1,000,000 文字以内です。1,000,000 文字を超えるデータが渡されると、 `ArgumentOutOfRange` の例外がスローされます。|
|`options`| オブジェクト| &lt;オプション&gt;|以下のプロパティのうち 1 つ以上を含むオブジェクト リテラル。|
|`options.asyncContext`| オブジェクト| &lt;オプション&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`options.coercionType`| [Office.CoercionType](office.md#coerciontype-string)| &lt;オプション&gt;| `text` の場合、Outlook Web アプリ と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。<br/><br/> `html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web アプリ では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、 `InvalidDataFormat` エラーが返されます。<br/><br/> `coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。|
|`callback`| 関数||メソッドが完了すると、`callback`パラメータで渡された関数が、シングル パラメータ、`asyncResult`、で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。 |

##### <a name="requirements"></a>必要条件

|要件| 値|
|---|---|
|[メールボックスに必要な設定バージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.2以降|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成|

##### <a name="example"></a>例

```JavaScript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```