---
title: Office.context.mailbox - 要件セット 1.1
description: Outlookメールボックス API 要件セット 1.1 バージョンのメールボックス オブジェクト モデル。
ms.date: 03/18/2020
ms.localizationpriority: medium
ms.openlocfilehash: 2bf509a790718a83e4ce1c6284904dd0887f9049
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151309"
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
| [診断](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1&preserve-view=true#diagnostics) | ReadItem | 作成<br>読み取り | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1&preserve-view=true#ewsUrl) | ReadItem | 作成<br>読み取り | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restricted | 作成<br>読み取り | [アイテム](/javascript/api/outlook/office.item?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1&preserve-view=true#userProfile) | ReadItem | 作成<br>読み取り | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.1&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>メソッド

| メソッド | 最小値<br>アクセス許可レベル | モード | 最小値<br>要件セット |
|---|---|---|:---:|
| [convertToLocalClientTime(timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1&preserve-view=true#convertToLocalClientTime_timeValue_) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToUtcClientTime(input)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1&preserve-view=true#convertToUtcClientTime_input_) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1&preserve-view=true#displayAppointmentForm_itemId_) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1&preserve-view=true#displayMessageForm_itemId_) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1&preserve-view=true#displayNewAppointmentForm_parameters_) | ReadItem | 読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1&preserve-view=true#getCallbackTokenAsync_callback__userContext_) | ReadItem | 作成<br>読み取り | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1&preserve-view=true#getUserIdentityTokenAsync_callback__userContext_) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1&preserve-view=true#makeEwsRequestAsync_data__callback__userContext_) | ReadWriteMailbox | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
