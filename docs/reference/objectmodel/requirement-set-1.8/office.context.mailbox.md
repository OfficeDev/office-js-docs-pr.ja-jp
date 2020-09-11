---
title: Office. メールボックス要件セット1.8
description: Outlook Mailbox API 要件セット1.8 バージョンのメールボックスオブジェクトモデル。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 673c40653b813a1c0d4a3faf3cb53877ccac00db
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431515"
---
# <a name="mailbox-requirement-set-18"></a>メールボックス (要件セット 1.8)

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
| [ダン](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#diagnostics) | ReadItem | 作成<br>読み取り | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.8&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#ewsurl) | ReadItem | 作成<br>読み取り | 文字列 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restricted | 作成<br>読み取り | [アイテム](/javascript/api/outlook/office.item?view=outlook-js-1.8&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [masterCategories](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#mastercategories) | ReadWriteMailbox | 作成<br>読み取り | [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#resturl) | ReadItem | 作成<br>読み取り | 文字列 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#userprofile) | ReadItem | 作成<br>読み取り | [プロファイル](/javascript/api/outlook/office.userprofile?view=outlook-js-1.8&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>メソッド

| Method | 最小値<br>アクセス許可レベル | モード | 最小値<br>要件セット |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | ReadItem | 作成<br>読み取り | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId (itemId、Office.mailboxenums.restversion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#converttoewsid-itemid--restversion-) | Restricted | 作成<br>読み取り | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime (timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#converttolocalclienttime-timevalue-) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId (itemId、Office.mailboxenums.restversion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#converttorestid-itemid--restversion-) | Restricted | 作成<br>読み取り | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime (input)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#converttoutcclienttime-input-) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#displayappointmentform-itemid-) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#displaymessageform-itemid-) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#displaynewappointmentform-parameters-) | ReadItem | 読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewMessageForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#displaynewmessageform-parameters-) | ReadItem | 作成<br>読み取り | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#getcallbacktokenasync-options--callback-) | ReadItem | 作成<br>読み取り | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#getcallbacktokenasync-callback--usercontext-) | ReadItem | 作成<br>読み取り | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#getuseridentitytokenasync-callback--usercontext-) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#removehandlerasync-eventtype--options--callback-) | ReadItem | 作成<br>読み取り | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>イベント

[Addハンドラ async](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-)と[removeハンドラ async](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true#removehandlerasync-eventtype--options--callback-)を使用して、次のイベントにサブスクライブし、サブスクライブを解除することができます。

| イベント | 説明 | 最小値<br>要件セット |
|---|---|:---:|
|`ItemChanged`| 作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
