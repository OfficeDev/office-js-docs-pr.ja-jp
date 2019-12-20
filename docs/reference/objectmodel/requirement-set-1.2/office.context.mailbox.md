---
title: Office. メールボックス要件セット1.2
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 99e755f995debc992001cf2549c46290c76f2c03
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814977"
---
# <a name="mailbox"></a>mailbox

### <a name="officeofficemdcontextofficecontextmdmailbox"></a>[Office](office.md)[.context](office.context.md).mailbox

Microsoft Outlook の Outlook アドイン オブジェクト モデルへのアクセスを提供します。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| 制限あり|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

## <a name="properties"></a>プロパティ

| プロパティ | 最小値<br>アクセス許可レベル | モード | 戻り値の種類 | 最小値<br>要件セット |
|---|---|---|---|:---:|
| [ダン](office.context.mailbox.diagnostics.md) | ReadItem | 作成<br>読み取り | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.2) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2#ewsurl) | ReadItem | 作成<br>読み取り | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | 制限あり | 作成<br>読み取り | [アイテム](/javascript/api/outlook/office.item?view=outlook-js-1.2) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [userProfile](office.context.mailbox.userProfile.md) | ReadItem | 作成<br>読み取り | [プロファイル](/javascript/api/outlook/office.userprofile?view=outlook-js-1.2) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>メソッド

| メソッド | 最小値<br>アクセス許可レベル | モード | 最小値<br>要件セット |
|---|---|---|:---:|
| [convertToLocalClientTime](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2#converttolocalclienttime-timevalue-) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToUtcClientTime](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2#converttoutcclienttime-input-) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2#displayappointmentform-itemid-) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2#displaymessageform-itemid-) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2#displaynewappointmentform-parameters-) | ReadItem | 読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getCallbackTokenAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2#getcallbacktokenasync-callback--usercontext-) | ReadItem | 読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2#getuseridentitytokenasync-callback--usercontext-) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
