
# <a name="item"></a>項目

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| 制限あり|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|

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

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook13officeattachmentdetails"></a>添付ファイル：配列.<[添付ファイルの詳細 ](/javascript/api/outlook_1_3/office.attachmentdetails)>

項目の添付ファイルの配列を取得します。閲覧モード専用です。

> [!NOTE]
> 潜在的なセキュリティ問題により特定の種類のファイルは、Outlookでブロックされ、したがって戻ってきません。 詳細については、[「Outlook でブロックされた添付ファイル」](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)をご覧下さい。

##### <a name="type"></a>型:

*   配列.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)>

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

次のコードでは、現在の項目にあるすべての添付ファイルの詳細を含む HTML 文字列を作成します。

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

####  <a name="bcc-recipientsjavascriptapioutlook13officerecipients"></a>bcc:[受信者](/javascript/api/outlook_1_3/office.recipients)

メッセージの BCC (ブラインド カーボン コピー) 列上の 受信者を取得または更新するメソッドを提供するオブジェクトを取得します。 新規作成モードのみ。

##### <a name="type"></a>型:

*   [受信者](/javascript/api/outlook_1_3/office.recipients)

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成|

##### <a name="example"></a>例

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook13officebody"></a>本文:[本文](/javascript/api/outlook_1_3/office.body)

アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。

##### <a name="type"></a>型:

*   [本文](/javascript/api/outlook_1_3/office.body)

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a>cc: 配列。 <[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_3/office.recipients)

メッセージの CC (カーボン コピー) 受信者へのアクセスを提供します。 オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。

##### <a name="read-mode"></a>閲覧モード

`cc`プロパティは、メッセージの**CC**列にある各受信者一覧の`EmailAddressDetails`オブジェクトを含む配列を返します。コレクションは最大100個のメンバーに制限されています。

##### <a name="compose-mode"></a>新規作成モード

`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。

##### <a name="type"></a>型:

*   配列 。<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_3/office.recipients)

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|

##### <a name="example"></a>例

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a>（空白が可能）conversationId：文字列

特定のメッセージが含まれている電子メールの会話の識別子を取得します。

メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。

作成フォームの新しいアイテムに対してこのプロパティの Null を取得します。ユーザーが件名を設定し項目を保存する場合、`conversationId`プロパティは値を返します。

##### <a name="type"></a>型:

*   文字列

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|

#### <a name="datetimecreated-date"></a>dateTimeCreated: 日付

アイテムが作成された日時を取得します。閲覧モード専用です。

##### <a name="type"></a>型:

*   日付

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a>dateTimeModified: 日付

アイテムが最後に変更された日時を取得します。読み取り専用です。

> [!NOTE]
> このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。

##### <a name="type"></a>型:

*   日付

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook13officetime"></a>end:日付|[時間](/javascript/api/outlook_1_3/office.time)

予定が終了する日時を取得または設定します。

`end`プロパティは、協定世界時 (UTC) 形式の時刻値として表示されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。

##### <a name="read-mode"></a>閲覧モード

`end`プロパティは`Date`オブジェクトを返します。

##### <a name="compose-mode"></a>新規作成モード

`end`プロパティは`Time`オブジェクトを返します。

[ `Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。

##### <a name="type"></a>型:

*   日付| [時間](/javascript/api/outlook_1_3/office.time)

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|

##### <a name="example"></a>例

次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a>from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)

メッセージの送信者の電子メール アドレスを取得します。読み取り専用です。

メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) プロパティは同一人物を表します。その場合、`from` プロパティは委任者を、送信者プロパティは代理人を表します。

> [!NOTE]
> `from` プロパティ内の`EmailAddressDetails` オブジェクトの`recipientType`  プロパティは、 `undefined`です。

##### <a name="type"></a>型:

*   [EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

#### <a name="internetmessageid-string"></a>internetMessageId: 文字列

電子メール メッセージのインターネット メッセージ 識別子を取得します。読み取り専用です。

##### <a name="type"></a>型:

*   文字列

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a>itemClass: 文字列

選択された項目の Exchange Web サービスの項目 クラスを取得します。閲覧モード専用です。

`itemClass` プロパティは、選択したアイテムのメッセージ クラスを指定します。次は、メッセージまたは予定アイテムの既定のメッセージ クラスを示しています。

| 種類 | 説明 | 項目のクラス |
| --- | --- | --- |
| 予定アイテム | これらは、アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurence` の予定表アイテムです。 | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| メッセージ アイテム | これには、基本のメッセージ クラス として `IPM.Schedule.Meeting`  を使用する、既定のメッセージ クラス `IPM.Note`  会議出席依頼、返信および取り消しを持つ電子メール メッセージが含まれます。 | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

既定のメッセージ クラスを拡張したカスタム メッセージ クラス 、たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など)を作成できます。

##### <a name="type"></a>型:

*   文字列

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a>（空白が可能） itemId ：文字列

現在の項目の Exchange Web サービスのアイテム識別子を取得します。閲覧モード専用です。

> [!NOTE]
> `itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。 `itemId` プロパティは、Outlook の Entry ID または Outlook の REST API によって使用される ID と同じではありません。 この値を使用して REST API の呼び出しを行う前に [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)を使用して変換する必要があります。 詳細については、 [Outlook アドインから Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)を参照してください。

新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。

##### <a name="type"></a>型:

*   文字列

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

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

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype"></a>itemType:[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)

インスタンスが表しているアイテムの種類を取得します。

`itemType`プロパティは、`ItemType`列挙値の 1 つを返します。これは`item`オブジェクト インスタンスがメッセージまたは予定のどちらであるかを示すものです。

##### <a name="type"></a>型:

*   [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|

##### <a name="example"></a>例

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook13officelocation"></a>位置: 文字列|[](/javascript/api/outlook_1_3/office.location)位置

予定の場所を取得または設定します。

##### <a name="read-mode"></a>閲覧モード

`location` プロパティは、予定の場所を含む文字列を返します。

##### <a name="compose-mode"></a>新規作成モード

`location` プロパティは、予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。

##### <a name="type"></a>型:

*   文字列 | [場所](/javascript/api/outlook_1_3/office.location)

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|

##### <a name="example"></a>例

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a>normalizedSubject: 文字列

すべてのプレフィックス (`RE:` や `FWD:` を含む) が削除された項目の件名を取得します。閲覧モード専用です。

normalizedSubject プロパティは、電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたアイテムの件名を取得します。プレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject) プロパティを使用します。

##### <a name="type"></a>型:

*   文字列

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages"></a>notificationMessages:[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)

項目の通知メッセージを取得します。

##### <a name="type"></a>型:

*   [NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a>optionalAttendees: 配列.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_3/office.recipients)

イベントの任意の出席者へのアクセスを提供します。 オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。

##### <a name="read-mode"></a>閲覧モード

`optionalAttendees`プロパティは、会議への各任意出席者の`EmailAddressDetails`オブジェクトを含む配列を返します。

##### <a name="compose-mode"></a>新規作成モード

`optionalAttendees`プロパティは会議への任意出席者を取得または設定するためのメソッドを提供する`Recipients`オブジェクトを返します。

##### <a name="type"></a>型:

*   配列 。<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_3/office.recipients)

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|

##### <a name="example"></a>例

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a>開催者:[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)

指定の会議の開催者の電子メール アドレスを取得します。読み取り専用です。

##### <a name="type"></a>型:

*   [EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a>requiredAttendees: 配列 。<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_3/office.recipients)。

イベントの必須出席者へのアクセスを提供します。 オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。

##### <a name="read-mode"></a>閲覧モード

`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを含む配列を返します。

##### <a name="compose-mode"></a>新規作成モード

`requiredAttendees` プロパティは会議への必須出席者を取得または設定するためのメソッドを提供する `Recipients` オブジェクトを返します。

##### <a name="type"></a>型:

*   配列 。<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_3/office.recipients)

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|

##### <a name="example"></a>例

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a>送信者:[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)

電子メール送信者のメールアドレスを取得します。閲覧モード専用です。

メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。その場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。

> [!NOTE]
> `sender`プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType`プロパティは、`undefined`です。

##### <a name="type"></a>型:

*   [EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook13officetime"></a>開始: 日付 |[ 時間](/javascript/api/outlook_1_3/office.time)

予定を開始する日時を取得または設定します。

`start`プロパティは、協定世界時 (UTC) 形式の時刻値として表示されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime)メソッドを使用して、値をクライアントのローカル日時に変換することができます。

##### <a name="read-mode"></a>閲覧モード

`start`プロパティは`Date`オブジェクトを返します。

##### <a name="compose-mode"></a>新規作成モード

`start` プロパティは `Time` オブジェクトを返します。

[ `Time.setAsync` ](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。

##### <a name="type"></a>型:

*   日付| [時間](/javascript/api/outlook_1_3/office.time)

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|

##### <a name="example"></a>例

次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。

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

####  <a name="subject-stringsubjectjavascriptapioutlook13officesubject"></a>件名: 文字列 | [件名](/javascript/api/outlook_1_3/office.subject)

アイテムの件名フィールドに表示される説明を取得または設定します。

`subject`プロパティは、電子メールサーバーから送信された項目の全件名を取得または設定します。

##### <a name="read-mode"></a>閲覧モード

`subject`プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string)プロパティを使用して、行間にある`RE:`や`FW:`のなどのすべてのプレフィックスを削除した件名を取得します。

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a>新規作成モード

`subject`プロパティは、件名を取得または設定するためのメソッドを提供する`Subject`オブジェクトを返します。

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a>型:

*   文字列 | [件名](/javascript/api/outlook_1_3/office.subject)

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a>to: 配列。 [EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_3/office.recipients)

メッセージの **宛先**列にある受信者へのアクセスを提供します。 オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。

##### <a name="read-mode"></a>閲覧モード

`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。

##### <a name="compose-mode"></a>新規作成モード

`to` プロパティは、メッセージの **To** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。

##### <a name="type"></a>型:

*   配列 。<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_3/office.recipients)

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|

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

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a>addFileAttachmentAsync(uri, attachmentName, [オプション], [コールバック])

ファイルを添付ファイルとしてメッセージまたは予定に追加します。

`addFileAttachmentAsync`メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内の項目に添付します。

その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 属性| 説明|
|---|---|---|---|
|`uri`| 文字列||メッセージまたは予定に添付するファイルの場所を示す URIです。 2048 文字以内で入力してください。|
|`attachmentName`| 文字列||アップロード中に表示される添付ファイルがそのファイルの名前です。 255 文字以内で入力してください。|
|`options`| オブジェクト| &lt;任意&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。|
|`options.asyncContext`| オブジェクト| &lt;任意&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`| 関数| &lt;任意&gt;|メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。 <br/>成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。<br/>添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。|

##### <a name="errors"></a>エラー

| エラー コード | 説明 |
|------------|-------------|
| `AttachmentSizeExceeded` | 添付ファイルのサイズが上限を超えています。 |
| `FileTypeNotSupported` | 許可されていない拡張子付きの添付ファイルです。 |
| `NumberOfAttachmentsExceeded` | メッセージまたは予定の添付ファイルが多すぎます。 |

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成|

##### <a name="example"></a>例

```
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

メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。

`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つ項目を新規作成フォーム内の項目に添付します。コールバック メソッドを指定する場合、`asyncResult` というパラメータがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、または項目を添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメータを使用して、状態情報をコールバック メソッドに渡すことができます。

その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。

Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドは項目を、編集中以外の項目に添付できますが、これはサポートされておらず、推奨されていません。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 属性| 説明|
|---|---|---|---|
|`itemId`| 文字列||添付するアイテムの Exchange 識別子です。最大長は 100 文字です。|
|`attachmentName`| 文字列||添付するアイテムの件名です。最大長は 255 文字です。|
|`options`| オブジェクト| &lt;任意&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。|
|`options.asyncContext`| オブジェクト| &lt;任意&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`| 関数| &lt;任意&gt;|メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。 <br/>成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。<br/>添付ファイルの追加に失敗した場合、 エラーの説明を提供する`asyncResult` オブジェクトが `Error` オブジェクトに含まれます。|

##### <a name="errors"></a>エラー

| エラー コード | 説明 |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | メッセージまたは予定の添付ファイルが多すぎます。 |

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成|

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

####  <a name="close"></a>閉じる()

新規作成中の現在の項目を閉じます。

`close`メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。

> [!NOTE]
> Outlook on the webでは、項目が予定で、`saveAsync`を用いて事前に保存されている場合、項目が最後に保存されてから何も変更されていない場合でも、ユーザーに対して保存、破棄またはキャンセルするよう求めるプロンプトを表示します。

Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close`メソッドは無効になります。

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| 制限あり|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成|

#### <a name="displayreplyallformformdata"></a>displayReplyAllForm(formData)

選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。

> [!NOTE]
> このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。

Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。

文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。

`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 説明|
|---|---|---|
|`formData`| 文字列 |オブジェクト| |回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。<br/>**または**<br/>本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクトです。オブジェクトの定義は次のとおりです。 |
| `formData.htmlBody` | 文字列 | &lt;任意&gt; | 回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。
| `formData.attachments` | 配列。&lt;オブジェクト&gt; | &lt;任意&gt; | ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。 |
| `formData.attachments.type` | 文字列 | | 添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` でなければいけません。 |
| `formData.attachments.name` | 文字列 | | 添付ファイル名を含む文字列です。最大長は 255 文字です。|
| `formData.attachments.url` | 文字列 | | `type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URIです。 |
| `formData.attachments.itemId` | 文字列 | | `type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムIDです。最大長が 100 文字の文字列です。 |
| `callback` | 関数 | &lt;任意&gt; | メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである シングル パラメーター `asyncResult` で呼び出されます。 |

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

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

本文と添付ファイルを返信します。

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

選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む返信フォームが表示されます。

> [!NOTE]
> このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。

Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。

文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。

`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 説明|
|---|---|---|
|`formData`| 文字列 |オブジェクト| | 回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。<br/>**または**<br/>本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクトです。オブジェクトの定義は次のとおりです。 |
| `formData.htmlBody` | 文字列 | &lt;任意&gt; | 回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。
| `formData.attachments` | 配列。&lt;オブジェクト&gt; | &lt;任意&gt; | ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。 |
| `formData.attachments.type` | 文字列 | | 添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` でなければいけません。 |
| `formData.attachments.name` | 文字列 | | 添付ファイル名を含む文字列です。最大長は 255 文字です。|
| `formData.attachments.url` | 文字列 | | `type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URIです。 |
| `formData.attachments.itemId` | 文字列 | | `type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムIDです。最大長が 100 文字の文字列です。 |
| `callback` | 関数 | &lt;任意&gt; | メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである シングル パラメーター `asyncResult` で呼び出されます。 |

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

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

本文と添付ファイルを返信します。

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

#### <a name="getentities--entitiesjavascriptapioutlook13officeentities"></a>getEntities() → {[エンティティ](/javascript/api/outlook_1_3/office.entities)}

選択したアイテムの本文にあるエンティティを取得します。

> [!NOTE]
> このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="returns"></a>戻り値 :

種類: [エンティティ](/javascript/api/outlook_1_3/office.entities)

##### <a name="example"></a>例

次の例では、現在の項目の本文内にある連絡先のエンティティにアクセスします。

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a>getEntitiesByType(entityType)] → [(空白可能) {<(String|[連絡先](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[電話番号](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion)) >}

選択した項目で見つかった指定のエンティティ型のエンティティすべてを含む配列を取得します。

> [!NOTE]
> このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 説明|
|---|---|---|
|`entityType`| [Office.MailboxEnums.EntityType](/javascript/api/outlook_1_3/office.mailboxenums.entitytype)|EntityType 列挙値の 1 つです。|

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| 制限あり|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="returns"></a>戻り値 :

`entityType` に渡された値が有効な `EntityType` 列挙型のメンバーでない場合、メソッドは 空白を返します。 指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。 そうでない場合、返される配列内のオブジェクトの種類は、 `entityType` パラメータ内で要求されたエンティティの種類によって異なります。

このメソッドを使用する最小限のアクセス許可レベルは **制限あり** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。

| の値 `entityType` | 返される配列内のオブジェクトの型 | 必要なアクセス許可のレベル |
| --- | --- | --- |
| `Address` | 文字列 | **制限あり** |
| `Contact` | 連絡先 | **ReadItem** |
| `EmailAddress` | 文字列 | **ReadItem** |
| `MeetingSuggestion` | MeetingSuggestion | **ReadItem** |
| `PhoneNumber` | PhoneNumber | **制限あり** |
| `TaskSuggestion` | TaskSuggestion | **ReadItem** |
| `URL` | 文字列 | **制限あり** |

型:Array.<(文字列|[連絡先](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>

##### <a name="example"></a>例

次の例は、現在のアイテムの本文にある郵便アドレスを表す文字列の配列にアクセスする方法を示します。

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a>getFilteredEntitiesByName(name)] → [(Null 許容) {<(文字列| [連絡先](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[電話番号 ](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}

マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。

> [!NOTE]
> このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。

`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 説明|
|---|---|---|
|`name`| 文字列|一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前です。|

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="returns"></a>戻り値 :

`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。

型:Array.<(文字列|[連絡先](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>

#### <a name="getregexmatches--object"></a>getRegExMatches() → {オブジェクト。

選択した項目内で、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。

> [!NOTE]
> このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。

`getRegExMatches` メソッドは、マニフェスト XML ファイル内にある各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定された項目のプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。

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

`getRegExMatches` から返されるオブジェクトには、 `fruits` および `veggies` という 2 つのプロパティがありえます。

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

アイテムの 本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。代わりに [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="returns"></a>戻り値 :

マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクトです。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。

<dl class="param-type">

<dt>型</dt>

<dd>オブジェクト</dd>

</dl>

##### <a name="example"></a>例

次の例は、マニフェストで指定された正規表現の <rule> 要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a>getRegExMatchesByName(name)] → [(Null許容) {配列. < 文字列 >}

選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。

> [!NOTE]
> このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。

`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。

項目の 本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、項目の本文全体を返さないようにします。`.*` などの正規表現を使用して項目の本文全体を取得しても、期待する結果が返されないことがあります。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 説明|
|---|---|---|
|`name`| 文字列|一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前です。|

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="returns"></a>戻り値 :

マニフェスト XML ファイルで定義された正規表現に一致する文字列を含む配列です。

<dl class="param-type">

<dt>種類</dt>

<dd>配列. < 文字列 ></dd>

</dl>

##### <a name="example"></a>例

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a>getSelectedDataAsync(coercionType, [オプション], コールバック) → {文字列}

メッセージの件名または本文から非同期的に選択したデータを返します。

選択したデータがなく、カーソルが本文または件名にある場合、メソッドは選択したデータに対して Null を返します。本文または件名以外のフィールドが選択されている場合、メソッドは `InvalidSelection` エラーを返します。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 属性| 説明|
|---|---|---|---|
|`coercionType`| [Office.CoercionType](office.md#coerciontype-string)||データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。|
|`options`| オブジェクト| &lt;任意&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。|
|`options.asyncContext`| オブジェクト| &lt;任意&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`| 関数||メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。<br/><br/>コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data`を呼び出します。 選択元のソース プロパティにアクセスするには、 `asyncResult.value.sourceProperty` を呼び出します。これは <   または `body` `subject` になります。|

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.2|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成|

##### <a name="returns"></a>戻り値 :

`coercionType`で決定された書式設定の文字列として選択されたデータです。

<dl class="param-type">

<dt>種類</dt>

<dd>文字列</dd>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a>loadCustomPropertiesAsync(callback, [userContext])

選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。

カスタム プロパティは、アプリケーションごと、アイテムごとにキーと値のペアとして保管されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在の項目および現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、項目上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 属性| 説明|
|---|---|---|---|
|`callback`| 関数||メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。<br/><br/>カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) オブジェクトとして指定されます。 項目からカスタム プロパティを取得、設定、削除して、サーバーにカスタム プロパティのセット バックに対する変更を保存するのに、このオブジェクトが使用できます。|
|`userContext`| オブジェクト| &lt;任意&gt;|開発者は、コールバック 関数でアクセスしたいオブジェクトを提供できます。 このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。|

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0以降|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a>removeAttachmentAsync(attachmentId、[オプション]、 [コールバック])

メッセージまたは予定から添付ファイルを削除します。

`removeAttachmentAsync`メソッドは、指定した識別子の添付ファイルを項目 から削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web アプリ とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別のウィンドウで操作を継続すると、セッションは終了します。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 属性| 説明|
|---|---|---|---|
|`attachmentId`| 文字列||削除する添付ファイルの識別子です。文字列の最大長は 100 文字です。|
|`options`| オブジェクト| &lt;任意&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。|
|`options.asyncContext`| オブジェクト| &lt;任意&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`| 関数| &lt;任意&gt;|メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。 <br/>添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。|

##### <a name="errors"></a>エラー

| エラー コード | 説明 |
|------------|-------------|
| `InvalidAttachmentId` | 添付ファイルの識別子が存在しません。 |

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.1|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成|

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

####  <a name="saveasyncoptions-callback"></a>saveAsync ([オプション] 、コールバック)

アイテムを非同期的に保存します。

呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッド経由でアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。

> [!NOTE]
> アドインが、WS または REST API を使用しようとして`itemId`を取得するために、新規作成モードでアイテム上の`saveAsync`を呼び出す場合、Outlook キャッシュ モードでは、アイテムがサーバーと実際に同期するまでに時間がかかる場合があることに注意してください。 項目が同期されるまで、 `itemId` を使用すると、エラーが返されます。

予定はドラフト状態にはならないため、作成モードで予定に`saveAsync`が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。

> [!NOTE]
> 次のクライアントは、新規作成モードで予定上の `saveAsync` に対して様々なふるまいをします。
>
> - Mac Outlook は、作成モードの会議で`saveAsync`をサポートしていません。 Mac Outlookの会議場で  `saveAsync` を呼びだすと、エラーが返されます。
> - 作成モードの予定上で`saveAsync`が呼び出されると、Outlook on the web は常に、招待または更新を送信します。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 属性| 説明|
|---|---|---|---|
|`options`| オブジェクト| &lt;任意&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。|
|`options.asyncContext`| オブジェクト| &lt;任意&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`callback`| 関数||メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。<br/><br/>成功すると、アイテム識別子が`asyncResult.value`プロパティに提供されます。|

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成|

##### <a name="examples"></a>例

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

次の例は、コールバック関数に渡される `result` パラメータの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a>setSelectedDataAsync(data, [options], callback)

メッセージの本文または件名に非同期的にデータを挿入します。

`setSelectedDataAsync`メソッドは、指定された文字列を項目のサブジェクトまたは本文のカーソル位置に挿入するか、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。

##### <a name="parameters"></a>パラメータ :

|名前| 種類| 属性| 説明|
|---|---|---|---|
|`data`| 文字列||挿入されるデータです。データの長さは 1,000,000 文字以内です。1,000,000 文字を超えるデータが渡されると、 `ArgumentOutOfRange` の例外がスローされます。|
|`options`| オブジェクト| &lt;任意&gt;|次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。|
|`options.asyncContext`| オブジェクト| &lt;任意&gt;|開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。|
|`options.coercionType`| [Office.CoercionType](office.md#coerciontype-string)| &lt;任意&gt;|`text` の場合、Outlook Web アプリ と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。<br/><br/>`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。<br/><br/>`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。|
|`callback`| 関数||メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。 |

##### <a name="requirements"></a>要件

|必要条件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.2|
|[アクセス許可に必要なレベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成|

##### <a name="example"></a>例

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```