---
title: Office のメールボックス-プレビュー要件セット
description: Outlook メールボックス API プレビュー要件セットのバージョンのメールボックスオブジェクトモデル。
ms.date: 06/17/2020
localization_priority: Normal
ms.openlocfilehash: 2aafc817555a59dcdadac2cc34bb69e487439639
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778670"
---
# <a name="mailbox-preview-requirement-set"></a>メールボックス (プレビュー要件セット)

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
| [ダン](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#diagnostics) | ReadItem | 作成<br>Read | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-preview) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#ewsurl) | ReadItem | 作成<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restricted | 作成<br>Read | [アイテム](/javascript/api/outlook/office.item?view=outlook-js-preview) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [masterCategories](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#mastercategories) | ReadWriteMailbox | 作成<br>Read | [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-preview) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#resturl) | ReadItem | 作成<br>Read | String | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#userprofile) | ReadItem | 作成<br>Read | [プロファイル](/javascript/api/outlook/office.userprofile?view=outlook-js-preview) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>メソッド

| Method | 最小値<br>アクセス許可レベル | モード | 最小値<br>要件セット |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | ReadItem | 作成<br>Read | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId (itemId、Office.mailboxenums.restversion)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#converttoewsid-itemid--restversion-) | Restricted | 作成<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime (timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#converttolocalclienttime-timevalue-) | ReadItem | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId (itemId、Office.mailboxenums.restversion)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#converttorestid-itemid--restversion-) | Restricted | 作成<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime (input)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#converttoutcclienttime-input-) | ReadItem | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displayappointmentform-itemid-) | ReadItem | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentFormAsync (itemId, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displayappointmentform-itemid--options--callback-) | ReadItem | 作成<br>Read | [プレビュー](outlook-requirement-set-preview.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaymessageform-itemid-) | ReadItem | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageFormAsync (itemId, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaymessageform-itemid--options--callback-) | ReadItem | 作成<br>Read | [プレビュー](outlook-requirement-set-preview.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewappointmentform-parameters-) | ReadItem | Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentFormAsync (parameters, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewappointmentform-parameters--options--callback-) | ReadItem | Read | [プレビュー](outlook-requirement-set-preview.md) |
| [displayNewMessageForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewmessageform-parameters-) | ReadItem | 作成<br>Read | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [displayNewMessageFormAsync (parameters, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewmessageform-parameters--options--callback-) | ReadItem | 作成<br>Read | [プレビュー](outlook-requirement-set-preview.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#getcallbacktokenasync-options--callback-) | ReadItem | 作成<br>Read | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#getcallbacktokenasync-callback--usercontext-) | ReadItem | 作成<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#getuseridentitytokenasync-callback--usercontext-) | ReadItem | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | ReadItem | 作成<br>Read | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>イベント

[Addハンドラ async](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-)と[removeハンドラ async](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-)を使用して、次のイベントにサブスクライブし、サブスクライブを解除することができます。

| イベント | 説明 | 最小値<br>要件セット |
|---|---|:---:|
|`ItemChanged`| 作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
|`OfficeThemeChanged`| メールボックスの Office テーマが変更されました。 | [プレビュー](../preview-requirement-set/outlook-requirement-set-preview.md) |
