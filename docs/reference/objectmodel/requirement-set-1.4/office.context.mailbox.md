---
title: Office.context.mailbox - 要件セット 1.4
description: Outlookメールボックス API 要件セット 1.4 バージョンのメールボックス オブジェクト モデル。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 765fe61c8107c11c8a713dce332d3221b319ff4275e157678b6b4a47e4261122
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57094634"
---
# <a name="mailbox-requirement-set-14"></a>メールボックス (要件セット 1.4)

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
| [診断](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4&preserve-view=true#diagnostics) | ReadItem | 作成<br>Read | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.4&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4&preserve-view=true#ewsUrl) | ReadItem | 作成<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restricted | 作成<br>Read | [アイテム](/javascript/api/outlook/office.item?view=outlook-js-1.4&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4&preserve-view=true#userProfile) | ReadItem | 作成<br>Read | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.4&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>メソッド

| メソッド | 最小値<br>アクセス許可レベル | モード | 最小値<br>要件セット |
|---|---|---|:---:|
| [convertToEwsId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4&preserve-view=true#convertToEwsId_itemId__restVersion_) | Restricted | 作成<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime(timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4&preserve-view=true#convertToLocalClientTime_timeValue_) | ReadItem | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4&preserve-view=true#convertToRestId_itemId__restVersion_) | Restricted | 作成<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime(input)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4&preserve-view=true#convertToUtcClientTime_input_) | ReadItem | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4&preserve-view=true#displayAppointmentForm_itemId_) | ReadItem | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4&preserve-view=true#displayMessageForm_itemId_) | ReadItem | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4&preserve-view=true#displayNewAppointmentForm_parameters_) | ReadItem | Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4&preserve-view=true#getCallbackTokenAsync_callback__userContext_) | ReadItem | 作成<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4&preserve-view=true#getUserIdentityTokenAsync_callback__userContext_) | ReadItem | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4&preserve-view=true#makeEwsRequestAsync_data__callback__userContext_) | ReadWriteMailbox | 作成<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
