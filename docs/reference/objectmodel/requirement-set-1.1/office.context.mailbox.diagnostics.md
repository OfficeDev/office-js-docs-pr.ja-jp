---
title: Office.--の要件セット1.1
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: a1787cb00b5d373c2051d40ccc219b05c8bea4af
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40815026"
---
# <a name="diagnostics"></a>診断

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics

Outlook アドインに診断情報を提供します。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

## <a name="properties"></a>プロパティ

| プロパティ | 最小値<br>アクセス許可レベル | モード | 戻り値の種類 | 最小値<br>要件セット |
|---|---|---|---|:---:|
| [名](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.1#hostname) | ReadItem | 作成<br>読み取り | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [上 diagnostics.hostversion](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.1#hostversion) | ReadItem | 作成<br>読み取り | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [OWAView](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.1#owaview) | ReadItem | 作成<br>読み取り | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
