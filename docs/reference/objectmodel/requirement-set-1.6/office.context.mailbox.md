---
title: Office.context.mailbox - 要件セット 1.6
description: Outlookメールボックス API 要件セット 1.6 バージョンのメールボックス オブジェクト モデル。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: e31244cabe998f0814dd62c204cfb08e24794ace
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590849"
---
# <a name="mailbox-requirement-set-16"></a>メールボックス (要件セット 1.6)

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
| [診断](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#diagnostics) | ReadItem | 作成<br>Read | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#ewsurl) | ReadItem | 作成<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restricted | 作成<br>Read | [項目](/javascript/api/outlook/office.item?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#resturl) | ReadItem | 作成<br>Read | String | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#userprofile) | ReadItem | 作成<br>Read | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>メソッド

| メソッド | 最小値<br>アクセス許可レベル | モード | 最小値<br>要件セット |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | ReadItem | 作成<br>Read | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#converttoewsid-itemid--restversion-) | Restricted | 作成<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime(timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#converttolocalclienttime-timevalue-) | ReadItem | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#converttorestid-itemid--restversion-) | Restricted | 作成<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime(input)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#converttoutcclienttime-input-) | ReadItem | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#displayappointmentform-itemid-) | ReadItem | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#displaymessageform-itemid-) | ReadItem | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#displaynewappointmentform-parameters-) | ReadItem | Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewMessageForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#displaynewmessageform-parameters-) | ReadItem | Read | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#getcallbacktokenasync-options--callback-) | ReadItem | 作成<br>Read | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#getcallbacktokenasync-callback--usercontext-) | ReadItem | 作成<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#getuseridentitytokenasync-callback--usercontext-) | ReadItem | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#removehandlerasync-eventtype--options--callback-) | ReadItem | 作成<br>Read | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>イベント

[addHandlerAsync と removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-)をそれぞれ使用して、次のイベントをサブスクライブおよび[サブスクライブ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true#removehandlerasync-eventtype--options--callback-)解除できます。

> [!IMPORTANT]
> イベントは、作業ウィンドウの実装でのみ使用できます。

| イベント | 説明 | 最小値<br>要件セット |
|---|---|:---:|
|`ItemChanged`| 作業ウィンドウOutlook表示する場合は、別のアイテムが選択されています。 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
