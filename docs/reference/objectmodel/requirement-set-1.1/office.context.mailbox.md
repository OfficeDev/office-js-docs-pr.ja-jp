---
title: Office. メールボックス要件セット1.1
description: Outlook Mailbox API 要件セット1.1 バージョンのメールボックスオブジェクトモデル。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: cec83b0148b47b6f50ef92f90e9094cd75b18e2b
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609007"
---
# <a name="mailbox-requirement-set-11"></a>メールボックス (要件セット 1.1)

### <a name="officecontextmailbox"></a>[Office](office.md)[.context](office.context.md).mailbox

Microsoft Outlook の Outlook アドイン オブジェクト モデルへのアクセスを提供します。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[最小限のアクセス許可レベル](../../../outlook/understanding-outlook-add-in-permissions.md)| 制限あり|
|[適用可能な Outlook のモード](../../../outlook/outlook-add-ins-overview.md#extension-points)| 新規作成または閲覧|

## <a name="properties"></a>プロパティ

| プロパティ | 最小値<br>アクセス許可レベル | モード | 戻り値の種類 | 最小値<br>要件セット |
|---|---|---|---|:---:|
| [ダン](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1#diagnostics) | ReadItem | 作成<br>Read | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.1) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1#ewsurl) | ReadItem | 作成<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restricted | 作成<br>Read | [Item](/javascript/api/outlook/office.item?view=outlook-js-1.1) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1#userprofile) | ReadItem | 作成<br>Read | [プロファイル](/javascript/api/outlook/office.userprofile?view=outlook-js-1.1) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>メソッド

| Method | 最小値<br>アクセス許可レベル | モード | 最小値<br>要件セット |
|---|---|---|:---:|
| [convertToLocalClientTime (timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1#converttolocalclienttime-timevalue-) | ReadItem | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToUtcClientTime (input)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1#converttoutcclienttime-input-) | ReadItem | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1#displayappointmentform-itemid-) | ReadItem | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1#displaymessageform-itemid-) | ReadItem | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1#displaynewappointmentform-parameters-) | ReadItem | Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1#getcallbacktokenasync-callback--usercontext-) | ReadItem | 作成<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1#getuseridentitytokenasync-callback--usercontext-) | ReadItem | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
