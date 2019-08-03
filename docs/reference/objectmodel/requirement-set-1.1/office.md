---
title: Office 名前空間 - 要件セット 1.1
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 01e60e0b01c7ca98bdff2c99dfcbe8944f822b81
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064362"
---
# <a name="office"></a>Office

Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

### <a name="namespaces"></a>名前空間

[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.1): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。

### <a name="members"></a>メンバー

#### <a name="asyncresultstatus-string"></a>AsyncResultStatus: String

非同期呼び出しの結果を指定します。

##### <a name="type"></a>型

*   String

##### <a name="properties"></a>プロパティ:

|名前| 種類| 説明|
|---|---|---|
|`Succeeded`| String|呼び出しが成功しました。|
|`Failed`| String|呼び出しが失敗しました。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

#### <a name="coerciontype-string"></a>CoercionType: String

呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。

##### <a name="type"></a>型

*   String

##### <a name="properties"></a>プロパティ:

|名前| 種類| 説明|
|---|---|---|
|`Html`| String|HTML 形式で返されるデータを要求します。|
|`Text`| String|テキスト形式で返されるデータを要求します。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

#### <a name="sourceproperty-string"></a>SourceProperty: String

呼び出されたメソッドによって返されるデータのソースを指定します。

##### <a name="type"></a>型

*   String

##### <a name="properties"></a>プロパティ:

|名前| 種類| 説明|
|---|---|---|
|`Body`| String|データのソースは、メッセージの本文です。|
|`Subject`| String|データのソースは、メッセージの件名です。|

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|
