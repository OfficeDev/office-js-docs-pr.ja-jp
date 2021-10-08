---
title: Office.context.mailbox - 要件セット 1.10
description: Outlookメールボックス API 要件セット 1.10 バージョンの Mailbox オブジェクト モデル。
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 052598a1fc0d4797f75ed60ae8d48afcd7c367d6
ms.sourcegitcommit: efd0966f6400c8e685017ce0c8c016a2cbab0d5c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/08/2021
ms.locfileid: "60237505"
---
# <a name="mailbox-requirement-set-110"></a>メールボックス (要件セット 1.10)

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
| [診断](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#diagnostics) | ReadItem | 作成<br>読み取り | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.10&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#ewsUrl) | ReadItem | 作成<br>読み取り | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restricted | 作成<br>読み取り | [Item](/javascript/api/outlook/office.item?view=outlook-js-1.10&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [masterCategories](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#masterCategories) | ReadWriteMailbox | 作成<br>読み取り | [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.10&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#restUrl) | ReadItem | 作成<br>読み取り | String | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#userProfile) | ReadItem | 作成<br>読み取り | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.10&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>メソッド

| メソッド | 最小値<br>アクセス許可レベル | モード | 最小値<br>要件セット |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | ReadItem | 作成<br>読み取り | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#convertToEwsId_itemId__restVersion_) | Restricted | 作成<br>読み取り | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime(timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#convertToLocalClientTime_timeValue_) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#convertToRestId_itemId__restVersion_) | Restricted | 作成<br>読み取り | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime(input)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#convertToUtcClientTime_input_) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#displayAppointmentForm_itemId_) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentFormAsync(itemId, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#displayAppointmentFormAsync_itemId__options__callback_) | ReadItem | 作成<br>読み取り | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#displayMessageForm_itemId_) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageFormAsync(itemId, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#displayMessageFormAsync_itemId__options__callback_) | ReadItem | 作成<br>読み取り | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#displayNewAppointmentForm_parameters_) | ReadItem | 読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentFormAsync(parameters, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#displayNewAppointmentFormAsync_parameters__options__callback_) | ReadItem | 読み取り | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayNewMessageForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#displayNewMessageForm_parameters_) | ReadItem | 読み取り | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [displayNewMessageFormAsync(parameters, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#displayNewMessageFormAsync_parameters__options__callback_) | ReadItem | 読み取り | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#getCallbackTokenAsync_options__callback_) | ReadItem | 作成<br>読み取り | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#getCallbackTokenAsync_callback__userContext_) | ReadItem | 作成<br>読み取り | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#getUserIdentityTokenAsync_callback__userContext_) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#makeEwsRequestAsync_data__callback__userContext_) | ReadWriteMailbox | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#removeHandlerAsync_eventType__options__callback_) | ReadItem | 作成<br>読み取り | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>イベント

[addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_)と[removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#removeHandlerAsync_eventType__options__callback_)をそれぞれ使用して、次のイベントをサブスクライブおよびサブスクライブ解除します。

> [!IMPORTANT]
> イベントは、作業ウィンドウの実装でのみ使用できます。

| [イベント](/javascript/api/office/office.eventtype?view=outlook-js-1.10&preserve-view=true) | 説明 | 最小値<br>要件セット |
|---|---|:---:|
|`ItemChanged`| 作業ウィンドウOutlook表示する場合は、別のアイテムが選択されています。 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
