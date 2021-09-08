---
title: Office名前空間 - 要件セット 1.9
description: Office API 要件セット 1.9 を使用Outlookアドインで使用できる名前空間メンバーを指定します。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 203b901c619e19a8e5b9255e36274e2f6e1d1658
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937408"
---
# <a name="office-mailbox-requirement-set-19"></a>Office (メールボックス要件セット 1.9)

Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[適用可能な Outlook のモード](../../../outlook/outlook-add-ins-overview.md#extension-points)| 新規作成または閲覧|

## <a name="properties"></a>プロパティ

| プロパティ | モード | 戻り値の種類 | 最小値<br>要件セット |
|---|---|---|:---:|
| [context](office.context.md) | 作成<br>読み取り | [Context](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a>列挙型

| 列挙体 | モード | 戻り値の種類 | 最小値<br>要件セット |
|---|---|---|:---:|
| [AsyncResultStatus](#asyncresultstatus-string) | 作成<br>読み取り | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [CoercionType](#coerciontype-string) | 作成<br>読み取り | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [EventType](#eventtype-string) | 作成<br>読み取り | String | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [SourceProperty](#sourceproperty-string) | 作成<br>読み取り | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a>名前空間

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.9&preserve-view=true): 、 など、Outlook固有の列挙の `ItemType` `EntityType` `AttachmentType` 数 `RecipientType` が `ResponseType` 含まれています `ItemNotificationMessageType` 。

## <a name="enumeration-details"></a>列挙の詳細

#### <a name="asyncresultstatus-string"></a>AsyncResultStatus: String

非同期呼び出しの結果を指定します。

##### <a name="type"></a>型

*   String

##### <a name="properties"></a>プロパティ

|名前| 型| 説明|
|---|---|---|
|`Succeeded`| String|呼び出しが成功しました。|
|`Failed`| String|呼び出しが失敗しました。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[適用可能な Outlook のモード](../../../outlook/outlook-add-ins-overview.md#extension-points)| 新規作成または閲覧|

<br>

---
---

#### <a name="coerciontype-string"></a>CoercionType: String

呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。

##### <a name="type"></a>型

*   String

##### <a name="properties"></a>プロパティ

|名前| 型| 説明|
|---|---|---|
|`Html`| String|HTML 形式で返されるデータを要求します。|
|`Text`| String|テキスト形式で返されるデータを要求します。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[適用可能な Outlook のモード](../../../outlook/outlook-add-ins-overview.md#extension-points)| 新規作成または閲覧|

<br>

---
---

#### <a name="eventtype-string"></a>EventType: String

イベント ハンドラーに関連付けられているイベントを指定します。

##### <a name="type"></a>型

*   String

##### <a name="properties"></a>プロパティ

| 名前 | 型 | 説明 | 最小要件セット |
|---|---|---|:---:|
|`AppointmentTimeChanged`| String | 選択した予定または系列の日付または時刻が変更されました。 | 1.7 |
|`AttachmentsChanged`| String | アイテムに添付ファイルが追加または削除されました。 | 1.8 |
|`EnhancedLocationsChanged`| String | 選択した予定の場所が変更されました。 | 1.8 |
|`ItemChanged`| String | 作業ウィンドウOutlook表示する場合は、別のアイテムが選択されています。 | 1.5 |
|`RecipientsChanged`| String | 選択したアイテムまたは予定の場所の受信者リストが変更されました。 | 1.7 |
|`RecurrenceChanged`| String | 選択した系列の定期的なパターンが変更されました。 | 1.7 |

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)| 1.5 |
|[適用可能な Outlook のモード](../../../outlook/outlook-add-ins-overview.md#extension-points)| 新規作成または閲覧|

<br>

---
---

#### <a name="sourceproperty-string"></a>SourceProperty: String

呼び出されたメソッドによって返されるデータのソースを指定します。

##### <a name="type"></a>型

*   String

##### <a name="properties"></a>プロパティ

|名前| 型| 説明|
|---|---|---|
|`Body`| String|データのソースは、メッセージの本文です。|
|`Subject`| String|データのソースは、メッセージの件名です。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[適用可能な Outlook のモード](../../../outlook/outlook-add-ins-overview.md#extension-points)| 新規作成または閲覧|
