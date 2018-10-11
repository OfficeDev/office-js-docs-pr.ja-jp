
# <a name="mailbox"></a>メールボックス

### [Office](Office.md)[.context](Office.context.md). mailbox

Microsoft Outlook と Microsoft Outlook on the web の Outlook アドイン オブジェクト モデルへのアクセスを提供します。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最低要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| 制限あり|
|[適用可能な Outlook モード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="members-and-methods"></a>メンバーとメソッド

| メンバー | 型 |
|--------|------|
| [ewsUrl](#ewsurl-string) | メンバー |
| [restUrl](#resturl-string) | メンバー |
| [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) | メソッド |
| [convertToEwsId](#converttoewsiditemid-restversion--string) | メソッド |
| [convertToLocalClientTime](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) | メソッド |
| [convertToRestId](#converttorestiditemid-restversion--string) | メソッド |
| [convertToUtcClientTime](#converttoutcclienttimeinput--date) | メソッド |
| [displayAppointmentForm](#displayappointmentformitemid) | メソッド |
| [displayMessageForm](#displaymessageformitemid) | メソッド |
| [displayNewAppointmentForm](#displaynewappointmentformparameters) | メソッド |
| [displayNewMessageForm](#displaynewmessageformparameters) | メソッド |
| [getCallbackTokenAsync](#getcallbacktokenasyncoptions-callback) | メソッド |
| [getCallbackTokenAsync](#getcallbacktokenasynccallback-usercontext) | メソッド |
| [getUserIdentityTokenAsync](#getuseridentitytokenasynccallback-usercontext) | メソッド |
| [makeEwsRequestAsync](#makeewsrequestasyncdata-callback-usercontext) | メソッド |

### <a name="namespaces"></a>名前空間

[diagnostics](Office.context.mailbox.diagnostics.md):Outlook アドインに診断情報を提供します。

[item](Office.context.mailbox.item.md):Outlook アドインのメッセージや予定にアクセスするためのメソッドとプロパティを提供します。

[userProfile](Office.context.mailbox.userProfile.md):Outlook アドインのユーザーに関する情報を提供します。

### <a name="members"></a>メンバー

#### <a name="ewsurl-string"></a>ewsUrl: 文字列

このメール アカウントの Exchange Web サービス (EWS) エンドポイントの URL を取得します。閲覧モードのみです。

> [!NOTE]
> このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。

`ewsUrl`値は、リモート サービスで、ユーザーのメールボックスに EWS 呼び出しを行うために使うことができます。たとえば、[選択したアイテムから添付ファイルを取得する](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。

閲覧モードでの`ewsUrl`メンバーを呼び出すには、アプリのマニフェスト内に指定されている** ReadItem **アクセス許可を有している必要があります。

新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)メソッドを呼び出してから、`ewsUrl`メンバーを使用する必要があります。アプリには、`saveAsync`メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。

##### <a name="type"></a>型:

*   文字列

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最低要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|

#### <a name="resturl-string"></a>restUrl: 文字列

この電子メール アカウントの REST エンドポイントの URL を取得します。

`restUrl`値は、ユーザーのメールボックスに [REST API](https://docs.microsoft.com/outlook/rest/) 呼び出しを行うために使用できます。

閲覧モードの`restUrl`メンバーを呼び出すには、アプリのマニフェスト内に指定されている** ReadItem **アクセス許可を有している必要があります。

新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback)メソッドを呼び出してから、`restUrl`メンバーを使用する必要があります。アプリには、`saveAsync`メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。

##### <a name="type"></a>型:

*   文字列

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最低要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

### <a name="methods"></a>メソッド

####  <a name="addhandlerasynceventtype-handler-options-callback"></a>addHandlerAsync (eventType、ハンドラ、 [ オプション ]、[ コールバック ])

サポートされているイベントのイベント ハンドラを追加します。

現在、サポートされているイベントの種類は、`Office.EventType.ItemChanged`と`Office.EventType.OfficeThemeChanged`。

##### <a name="parameters"></a>パラメータ:

| 名前: | 型 | 属性 | 説明 |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || ハンドラを呼び出す必要のあるイベント。 |
| `handler` | 関数 || イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメータを受け入れる必要があります。パラメータの`type`プロパティは、`eventType`に渡される`addHandlerAsync`パラメータと一致します。 |
| `options` | オブジェクト | &lt;任意&gt; | 次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。 |
| `options.asyncContext` | オブジェクト | &lt;任意&gt; | 開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。 |
| `callback` | 関数| &lt;任意&gt;|メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最低要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="example"></a>例

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="converttoewsiditemid-restversion--string"></a>convertToRestId( アイテム Id、 restVersion) → { 文字列 }

REST 形式のアイテム ID を EWS 形式に変換します。

> [!NOTE]
> このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。

REST API([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](http://graph.microsoft.io/) など  ) で取得されたアイテム ID は、Exchange Web サービス (EWS) に使用される形式とは異なる形式を使用します。`convertToEwsId`メソッドは、REST 形式の ID を EWS 用の適切な形式に変換します。

##### <a name="parameters"></a>パラメータ:

|名前:| 型| 説明|
|---|---|---|
|`itemId`| 文字列|Outlook REST API 形式のアイテム ID|
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion)|アイテム ID の取得に使用された Outlook REST API のバージョンを示す値。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最低要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| 制限あり|
|[適用可能な Outlook モード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="returns"></a>戻り値 :

型 : String

##### <a name="example"></a>例

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a>convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}

クライアントのローカル時間で時間情報が含まれている辞書を取得します。

Outlook 用メール アプリや Outlook Web App で使う日付と時刻は、異なるタイム ゾーンを使うことができます。Outlook では、クライアント コンピューターのタイム ゾーンを使います。Outlook Web App では、Exchange 管理センター (EAC) で設定されたタイム ゾーンを使います。ユーザー インターフェイスに表示される値が、常にユーザーが期待するタイム ゾーンと一致するように日付と時刻の値を処理する必要があります。

Outlook でメール アプリが実行されている場合、`convertToLocalClientTime`メソッドは、クライアント コンピューターのタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。Outlook Web Apps でメール アプリが実行されている場合、`convertToLocalClientTime`メソッドは、EAC に指定されたタイム ゾーンに設定された値のディクショナリ オブジェクトを返します。

##### <a name="parameters"></a>パラメータ:

|名前:| 型| 説明|
|---|---|---|
|`timeValue`| 日付|日付オブジェクト|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最低要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="returns"></a>戻り値:

型:[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)

####  <a name="converttorestiditemid-restversion--string"></a>convertToRestId (アイテム Id、restVersion)] → [{文字列}

EWS 形式のアイテム ID を REST 形式に変換します。

> [!NOTE]
> このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。

EWS または `itemId` プロパティで取得されるアイテム ID は、REST API ([Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) や [Microsoft Graph](http://graph.microsoft.io/) など) に使用される形式とは異なる形式を使用します。`convertToRestId` メソッドは、EWS 形式の ID を REST 用の適切な形式に変換します。

##### <a name="parameters"></a>パラメータ:

|名前:| 型| 説明|
|---|---|---|
|`itemId`| 文字列|文字列|
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion)|変換後の ID を使用する Outlook REST API のバージョンを示す値。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最低要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| 制限あり|
|[適用可能な Outlook モード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="returns"></a>戻り値 :

型 : String

##### <a name="example"></a>例

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a>convertToUtcClientTime(input) → {Date}

時間情報が含まれているディクショナリから日付オブジェクトを取得します。

メソッドは、`convertToUtcClientTime`ローカルの日付と時刻を含むディクショナリを、ローカルの日付と時刻の正しい値を持つ日付オブジェクトに変換します。

##### <a name="parameters"></a>パラメータ:

|名前:| 型| 説明|
|---|---|---|
|`input`| [現地時間](/javascript/api/outlook/office.LocalClientTime)|変換するローカル時刻の値。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最低要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="returns"></a>戻り値:

時間が UTC で表現された日付オブジェクト。

<dl class="param-type">

<dt>型</dt>

<dd>日付</dd>

</dl>

####  <a name="displayappointmentformitemid"></a>displayAppointmentForm(itemId)

既存の予定を表示します。

> [!NOTE]
> このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。

`displayAppointmentForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスで既存の予定を開きます。

Outlook for Mac では、この方法を使用して、定期的に繰り返される予定の一部ではない単発の予定、または定期的に繰り替えされる元の予定を表示できます。ただし、一連の予定のインスタンスは表示できません。これは、Outlook for Mac では、定期的に繰り返されるインスタンスのプロパティ  (アイテム ID を含む) にアクセスできないためです。

Outlook Web App では、このメソッドは、指定されたフォームの本文が 32 KB 以下の文字数の場合に、そのフォームを開きます。

指定のアイテム識別子が既存の予定を表していない場合は、クライアント コンピュータまたはデバイスで空白ウィンドウが開き、エラー メッセージは返されません。

##### <a name="parameters"></a>パラメータ:

|名前:| 型| 説明|
|---|---|---|
|`itemId`| 文字列|既存の予定の Exchange Web サービス (EWS) 識別子。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最低要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="example"></a>例

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a>displayMessageForm(itemId)

既存のメッセージを表示します。

> [!NOTE]
> このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。

`displayMessageForm` メソッドは、デスクトップ上の新しいウィンドウやモバイル デバイス上のダイアログ ボックスで既存のメッセージを開きます。

Outlook Web App では、このメソッドは、指定されたフォームの本文が 32 KB 以下の文字数の場合に、そのフォームを開きます。

指定のアイテム識別子が既存のメッセージを指定していない場合、クライアント コンピュータでメッセージは表示されず、エラー メッセージも返されません。

予定を表す `displayMessageForm` を含む `itemId` を使用しないでください。既存の予定を表示するには、`displayAppointmentForm` メソッドを使用し、新しい予定を作成するフォームを表示するには、`displayNewAppointmentForm` メソッドを使用します。

##### <a name="parameters"></a>パラメータ:

|名前:| 型| 説明|
|---|---|---|
|`itemId`| 文字列|既存のメッセージの Exchange Web サービス (EWS) 識別子。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最低要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="example"></a>例

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a>displayNewAppointmentForm(parameters)

新しい予定を作成するためのフォームを表示します。

> [!NOTE]
> このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。

`displayNewAppointmentForm` メソッドを使用すると、ユーザーが新しい予定または会議を作成できるフォームが開きます。パラメータを指定すると、予定のフォーム フィールドにパラメータの内容が自動的に設定されます。

Outlook Web App と OWA for Devices では、このメソッドは、出席者フィールドが含まれるフォームを必ず表示します。入力引数として出席者を指定しない場合、このメソッドにより **[ 保存 ]** ボタンのあるフォームが表示されます。出席者を指定した場合には、フォームにその出席者と **[ 送信 ]** ボタンが表示されます。

Outlook リッチ クライアントと Outlook RT で、`requiredAttendees`、`optionalAttendees`、または `resources` パラメータに出席者またはリソースを指定し、このメソッドを実行すると、**[ 送信 ]** ボタンがある会議フォームが表示されます。受信者を指定せずにこのメソッドを実行すると、**[ 保存して閉じる ]** ボタンがある予定フォームが表示されます。

パラメータのいずれかが指定のサイズ制限を超える場合、または不明なパラメータ名が指定されている場合は、例外が反映されます。

##### <a name="parameters"></a>パラメータ:

> [!NOTE]
> すべてのパラメータは省略可能です。

|名前:| 型| 説明|
|---|---|---|
| `parameters` | オブジェクト | 新しい予定を記述するパラメータの辞書。 |
| `parameters.requiredAttendees` | 配列。&lt;文字列&gt; | 配列です。&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt; | 予定に必要な各出席者のメール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。 |
| `parameters.optionalAttendees` | 配列。&lt;文字列&gt; | 配列です。&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt; | 予定の各任意の出席者のメール アドレスを含む文字列の配列、または `EmailAddressDetails` オブジェクトを含む配列。配列の上限は 100 エントリです。 |
| `parameters.start` | 日付 | 予定の開始日時を指定する `Date` オブジェクト。 |
| `parameters.end` | 日付 | 予定の終了日時を指定する `Date` オブジェクト。 |
| `parameters.location` | 文字列 | 予定の場所を含む文字列。文字列は最大 255 文字に制限されます。 |
| `parameters.resources` | 配列。&lt; 文字列&gt; | 予定に必要なリソースを含む文字列の配列。配列の上限は 100 エントリです。 |
| `parameters.subject` | 文字列 | 予定の件名を含む文字列です。文字列は最大 255 文字に制限されます。 |
| `parameters.body` | 文字列 | 予定の本文。本文の内容は、最大サイズが 32 KB に制限されます。 |

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最低要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```
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

#### <a name="displaynewmessageformparameters"></a>displayNewMessageForm(parameters)

新しいメッセージを作成するためのフォームを表示します。

`displayNewMessageForm` メソッドは、ユーザーが新しいメッセージを作成できるフォームを開きます。 パラメータを指定すると、メッセージ フォーム フィールドにはパラメータのコンテンツが自動的に入力されます。

パラメータのいずれかが指定のサイズ制限を超える場合、または不明なパラメータ名が指定されている場合は、例外が反映されます。

##### <a name="parameters"></a>パラメータ:

> [!NOTE]
> すべてのパラメータは省略可能です。

|名前:| 型| 説明|
|---|---|---|
| `parameters` | オブジェクト | 新しいメッセージを記述するパラメータの辞書。 |
| `parameters.toRecipients` | 配列。&lt;文字列&gt; | 配列です。&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt; | 電子メール アドレスを含む文字列の配列、または To 行の各受信者の `EmailAddressDetails` オブジェクトを含む配列。 配列の上限は 100 エントリです。 |
| `parameters.ccRecipients` | 配列。&lt;文字列&gt; | 配列です。&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt; | 電子メール アドレスを含む文字列の配列、または Cc 行の各受信者の `EmailAddressDetails` オブジェクトを含む配列。 配列の上限は 100 エントリです。 |
| `parameters.bccRecipients` | 配列。&lt;文字列&gt; | 配列です。&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt; | 電子メール アドレスを含む文字列の配列、または Bcc 行の各受信者の `EmailAddressDetails` オブジェクトを含む配列。 配列の上限は 100 エントリです。 |
| `parameters.subject` | 文字列 | メッセージの件名を含む文字列。 文字列は最大 255 文字に制限されます。 |
| `parameters.htmlBody` | 文字列 | メッセージの HTML 本文。 本文の内容は、最大サイズが 32 KB に制限されます。 |
| `parameters.attachments` | 配列。&lt;オブジェクト&gt; | 添付ファイルまたは添付アイテムである JSON オブジェクトの配列。 |
| `parameters.attachments.type` | 文字列 | 添付ファイルの種類を示します。添付ファイルの場合は `file`、添付アイテムの場合は `item` でなければなりません。 |
| `parameters.attachments.name` | 文字列 | 添付ファイル名を含む文字列で、最長 255 文字です。|
| `parameters.attachments.url` | 文字列 | `type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URIです。 |
| `parameters.attachments.isInline` | ブール値 | `type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。 |
| `parameters.attachments.itemId` | 文字列 | `type` が `item` に設定されている場合にのみ使用されます。 新しいメッセージに添付する、既存の電子メールの EWS 項目の id です。 最長 100 文字の文字列です。 |


##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最低要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 読み取り|

##### <a name="example"></a>例

```
Office.context.mailbox.displayNewMessageForm(
  {
    toRecipients: Office.context.mailbox.item.to, // Copy the To line from current item
    ccRecipients: ['sam@contoso.com'],
    subject: 'Outlook add-ins are cool!',
    htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
    attachments: [
      {
        type: 'file',
        name: 'image.png',
        url: 'http://contoso.com/image.png',
        isInline: true
      }
    ]
  });
```

#### <a name="getcallbacktokenasyncoptions-callback"></a>getCallbackTokenAsync ([ オプション ]、コールバック )

REST API または Exchange Web サービスを呼び出すために使用するトークンを含む文字列を取得します。

`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。

> [!NOTE]
> 可能であれば、アドインでは Exchange Web サービスの代わりに REST API を使用することをお勧めします。 

**REST トークン**

REST トークンが要求された場合 (`options.isRest = true`)、結果のトークンは Exchange Web サービスの呼び出しを認証するために機能しません。アドインがマニフェストで [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) アクセス許可を指定していない限り、トークンの範囲は現在のアイテムとその添付ファイルへの読み取り専用アクセスに制限されます。`ReadWriteMailbox` アクセス許可が指定されている場合は、結果のトークンは、メールを送信する機能など、メール、カレンダー、連絡先への読み取り/書き込みアクセスを付与します。

アドインでは、`restUrl` プロパティを使用して、REST API 呼び出しを行うときに使用する正しい URL を決定する必要があります。

**EWS トークン**

EWS トークンが要求された場合 (`options.isRest = false`)、結果トークンは REST API 呼び出しを認証するために機能しません。トークンの範囲は、現在のアイテムへのアクセスに制限されます。

アドインでは、`ewsUrl` プロパティを使用して、EWS 呼び出しを行うときに使用する正しい URL を決定する必要があります。

##### <a name="parameters"></a>パラメータ:

|名前:| 型| 属性| 説明|
|---|---|---|---|
| `options` | オブジェクト | &lt;任意&gt; | 次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。 |
| `options.isRest` | ブール値 |  &lt;任意&gt; | 提供されたトークンを Outlook REST API または Exchange Web サービスに使用するかどうかを決定します。既定値は `false`です。 |
| `options.asyncContext` | オブジェクト |  &lt;任意&gt; | 非同期メソッドに渡される状態データです。 |
|`callback`| 関数||メソッドが完了すると、`callback` パラメータに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトであるシングル パラメータ `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最低要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成と閲覧|

##### <a name="example"></a>例

```js
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getcallbacktokenasynccallback-usercontext"></a>getCallbackTokenAsync ( コールバック、[ ユーザー コンテキスト ])

Exchange Server から添付ファイルやアイテムを取得するために使うトークンを含む文字列を取得します。

`getCallbackTokenAsync` メソッドは、ユーザーのメールボックスをホストする Exchange Server から不透明なトークンを取得する非同期の呼び出しを行います。コールバック トークンの有効期間は 5 分です。

トークンと予定の識別子またはアイテムの識別子をサードパーティ システムに渡すことができます。サードパーティ システムは、トークンをベアラー承認トークンとして使用し、Exchange Web サービス (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) または [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) 操作を呼び出して、添付ファイルまたはアイテムを返します。たとえば 、[選択したアイテムから添付ファイルを取得する](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item)ためにリモート サービスを作成できます。

閲覧モードの `getCallbackTokenAsync` メソッドを呼び出すには、お使いのアプリがそのマニフェストで指定されている** ReadItem** アクセス許可を有している必要があります。

新規作成モードでは、[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) メソッドを呼び出してアイテムの識別子を `getCallbackTokenAsync` メソッドに渡す必要があります。アプリには、`saveAsync` メソッドを呼び出す **ReadWriteItem** アクセス許可が必要です。

##### <a name="parameters"></a>パラメータ:

|名前:| 型| 属性| 説明|
|---|---|---|---|
|`callback`| 関数||メソッドが完了すると、`callback` パラメータに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトであるシングル パラメータ `AsyncResult` で呼び出されます。トークンは、`asyncResult.value` プロパティで文字列として提供されます。|
|`userContext`| オブジェクト| &lt;任意&gt;|非同期メソッドに渡される状態データです。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最低要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成と閲覧|

##### <a name="example"></a>例

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a>getUserIdentityTokenAsync (コールバック、[ ユーザー コンテキスト ])

ユーザーと Office アドインを識別するトークンを取得します。

`getUserIdentityTokenAsync`メソッドは、[アドインとユーザーをサード パーティのシステムで識別して認証](https://docs.microsoft.com/outlook/add-ins/authentication)することのできるトークンを返します。

##### <a name="parameters"></a>パラメータ:

|名前:| 型| 属性| 説明|
|---|---|---|---|
|`callback`| 関数||メソッドが完了すると、`callback` パラメータに渡された関数が、シングル パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。<br/><br/>トークンは、`asyncResult.value` プロパティで文字列として提供されます。|
|`userContext`| オブジェクト| &lt;任意&gt;|非同期メソッドに渡される状態データです。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最低要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="example"></a>例

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a>makeEwsRequestAsync (データ、コールバック、[ ユーザー コンテキスト ])

ユーザーのメールボックスをホストしている Exchange Server 上の Exchange Web サービス (EWS) のサービスに対して非同期の要求を行います。

> [!NOTE]
> このメソッドは、次のシナリオではサポートされていません。
> - Outlook for iOS または Outlook for Android で
> - アドインの読み込み時 Gmail のメールボックスに
> 
> これらの場合では、アドインは[ REST Api を使用する](https://docs.microsoft.com/outlook/add-ins/use-rest-api)代わりに、ユーザーのメールボックスにアクセスする必要があります。

`makeEwsRequestAsync` メソッドは、アドインの代わりに Exchange に EWS 要求を送信します。 サポートされている EWS 操作の一覧については、「[Outlook のアドインからの web サービスを呼び出す](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support)」を参照してください。

`makeEwsRequestAsync` メソッドで、フォルダー関連アイテムを要求することはできません。

XML 要求では UTF-8 エンコードを指定する必要があります。

```
<?xml version="1.0" encoding="utf-8"?>
```

`makeEwsRequestAsync` メソッドを使用するには、アドインが **ReadWriteMailbox** アクセス許可を有していなければなりません。** ReadWriteMailbox** アクセス許可の使い方と、`makeEwsRequestAsync` メソッドで呼び出せる EWS 操作については、「[ユーザーのメールボックスへのメール アドイン アクセスのアクセス許可を指定する](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)」を参照してください。

> [!NOTE]
> サーバー管理者は、クライアント アクセス サーバーの EWS ディレクトリで `OAuthAuthentication` を true に設定して、`makeEwsRequestAsync` メソッドで EWS 要求を行うことができるようにする必要があります。

##### <a name="version-differences"></a>バージョンの相違点

バージョン 15.0.4535.1004 より前のバージョンの Outlook で実行しているメール アプリで `makeEwsRequestAsync` メソッドを使う場合は、エンコード値を `ISO-8859-1` に設定する必要があります。

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

メール アプリが Web 上の Outlook で実行されているときに、エンコード値を設定する必要はありません。mailbox.diagnostics.hostNameプロパティを使用すると、メール アプリが Outlook または Web 上の Outlook で実行されているかどうかを判断できます。実行中の Outlook のバージョンは、mailbox.diagnostics.hostVersion プロパティを使用して確認できます。

##### <a name="parameters"></a>パラメータ:

|名前:| 型| 属性| 説明|
|---|---|---|---|
|`data`| 文字列||EWS 要求です。|
|`callback`| 関数||メソッドが完了すると、`callback` パラメータに渡された関数が、シングル パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。<br/><br/>EWS 呼び出しの XML 結果は、`asyncResult.value`プロパティ内の文字列として提供されています。 結果のサイズが 1 MB を超えている場合、エラー メッセージが返されます。|
|`userContext`| オブジェクト| &lt;任意&gt;|非同期メソッドに渡される状態データです。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最低要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteMailbox|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 作成または読み取り|

##### <a name="example"></a>例

次の例は、`makeEwsRequestAsync` を呼び出し、`GetItem` 操作を使用してアイテムの件名を取得します。

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