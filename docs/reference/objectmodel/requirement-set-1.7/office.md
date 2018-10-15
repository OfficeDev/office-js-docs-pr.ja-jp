 

# <a name="office"></a>Office

Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共有 API](/javascript/api/office)」を参照してください。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|

##### <a name="members-and-methods"></a>メンバーとメソッド

| メンバー | 型 |
|--------|------|
| [AsyncResultStatus](#asyncresultstatus-string) | メンバー |
| [CoercionType](#coerciontype-string) | メンバー |
| [EventType](#eventtype-string) | メンバー |
| [SourceProperty](#sourceproperty-string) | メンバー |

### <a name="namespaces"></a>名前空間

[context](office.context.md):Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。

[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype):ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。

### <a name="members"></a>メンバー

####  <a name="asyncresultstatus-string"></a>AsyncResultStatus: 文字列

非同期呼び出しの結果を指定します。

##### <a name="type"></a>型:

*   文字列

##### <a name="properties"></a>プロパティ:

|名前| 型| 説明|
|---|---|---|
|`Succeeded`| 文字列|呼び出しが成功しました。|
|`Failed`| 文字列|呼び出しが失敗しました。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|

---

####  <a name="coerciontype-string"></a>CoercionType: 文字列

呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。

##### <a name="type"></a>型:

*   文字列

##### <a name="properties"></a>プロパティ:

|名前| 型| 説明|
|---|---|---|
|`Html`| 文字列|HTML 形式で返されるデータを要求します。|
|`Text`| 文字列|テキスト形式で返されるデータを要求します。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|

---

####  <a name="eventtype-string"></a>イベントの種類: 文字列

イベント ハンドラに関連付けられているイベントを指定します。

##### <a name="type"></a>型:

*   文字列

##### <a name="properties"></a>プロパティ:

| 名前 | 型 | 説明 | 最小要件セット |
|---|---|---|---|
|`AppointmentTimeChanged`| 文字列 | 選択した予定または系列の日付または時間が変更されました。 | 1.7 |
|`ItemChanged`| 文字列 | 選択したアイテムが変更されました。 | 1.5 |
|`RecipientsChanged`| 文字列 | 選択したアイテムまたは予定の場所の受信者の一覧が変更されました。 | 1.7 |
|`RecurrenceChanged`| 文字列 | 選択した系列の定期的なパターンが変更されました。 | 1.7 |

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックス要件セットの最小バージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り |

---

####  <a name="sourceproperty-string"></a>SourceProperty: 文字列

呼び出されたメソッドによって返されるデータのソースを指定します。

##### <a name="type"></a>型:

*   文字列

##### <a name="properties"></a>プロパティ:

|名前| 型| 説明|
|---|---|---|
|`Body`| 文字列|データのソースは、メッセージの本文です。|
|`Subject`| 文字列|データのソースは、メッセージの件名です。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|