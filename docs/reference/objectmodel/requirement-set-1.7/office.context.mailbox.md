---
title: Office. メールボックス要件セット1.7
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 0572ff2ce3a21cc79bbb16a2ac1a9d0da86ac57b
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814599"
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
| [ダン](office.context.mailbox.diagnostics.md) | ReadItem | 作成<br>読み取り | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.7) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#ewsurl) | ReadItem | 作成<br>読み取り | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | 制限あり | 作成<br>読み取り | [アイテム](/javascript/api/outlook/office.item?view=outlook-js-1.7) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#resturl) | ReadItem | 作成<br>読み取り | String | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](office.context.mailbox.userProfile.md) | ReadItem | 作成<br>読み取り | [プロファイル](/javascript/api/outlook/office.userprofile?view=outlook-js-1.7) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>メソッド

| メソッド | 最小値<br>アクセス許可レベル | モード | 最小値<br>要件セット |
|---|---|---|:---:|
| [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#addhandlerasync-eventtype--handler--options--callback-) | ReadItem | 作成<br>読み取り | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#converttoewsid-itemid--restversion-) | 制限あり | 作成<br>読み取り | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#converttolocalclienttime-timevalue-) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#converttorestid-itemid--restversion-) | 制限あり | 作成<br>読み取り | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#converttoutcclienttime-input-) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#displayappointmentform-itemid-) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#displaymessageform-itemid-) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#displaynewappointmentform-parameters-) | ReadItem | 読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewMessageForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#displaynewmessageform-parameters-) | ReadItem | 作成<br>読み取り | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [getCallbackTokenAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#getcallbacktokenasync-options--callback-) | ReadItem | 作成<br>読み取り | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#getcallbacktokenasync-callback--usercontext-) | ReadItem | 作成<br>読み取り | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#getuseridentitytokenasync-callback--usercontext-) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#removehandlerasync-eventtype--options--callback-) | ReadItem | 作成<br>読み取り | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>イベント

[Addハンドラ async](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#addhandlerasync-eventtype--handler--options--callback-)と[removeハンドラ async](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7#removehandlerasync-eventtype--options--callback-)を使用して、次のイベントにサブスクライブし、サブスクライブを解除することができます。

| イベント | 説明 | 最小値<br>要件セット |
|---|---|:---:|
|`ItemChanged`| 作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
