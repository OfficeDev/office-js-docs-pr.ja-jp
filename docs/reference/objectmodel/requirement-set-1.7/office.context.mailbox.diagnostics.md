---
title: Office.--の要件セット1.7
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 2a79dbe7d392b809cf0de0b5ee7096473ea3e197
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127192"
---
# <a name="diagnostics"></a>診断

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a>[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics

Outlook アドインに診断情報を提供します。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

##### <a name="members-and-methods"></a>メンバーとメソッド

| メンバー | 種類 |
|--------|------|
| [名](#hostname-string) | Member |
| [上 diagnostics.hostversion](#hostversion-string) | Member |
| [OWAView](#owaview-string) | メンバー |

### <a name="members"></a>メンバー

#### <a name="hostname-string"></a>hostName: String

ホスト アプリケーションの名前を表す文字列を取得します。

文字列は、値 `Outlook`、`Mac Outlook`、`OutlookIOS`、または `OutlookWebApp` のいずれかになります。

##### <a name="type"></a>型

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

---
---

#### <a name="hostversion-string"></a>hostVersion: String

ホスト アプリケーションまたは Exchange Server のバージョンを表す文字列を取得します。

メールアドインが Outlook デスクトップクライアントまたは iOS で実行されている場合、 `hostVersion`このプロパティはホストアプリケーションのバージョン (outlook) を返します。 Web 上の Outlook では、このプロパティは Exchange サーバーのバージョンを返します。 たとえば、文字列 `15.0.468.0` です。

##### <a name="type"></a>型

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

---
---

#### <a name="owaview-string"></a>OWAView: String

Web 上の Outlook の現在のビューを表す文字列を取得します。

返される文字列は、値 `OneColumn`、`TwoColumns`、または `ThreeColumns` のいずれかになります。

ホストアプリケーションが web 上の Outlook ではない場合、このプロパティにアクセスする`undefined`と、になります。

Outlook on the web には、画面とウィンドウの幅、および表示できる列の数に対応する3つのビューがあります。

*   画面幅が狭い場合に表示される `OneColumn`。 Outlook on the web では、スマートフォンの画面全体でこのような単一の列のレイアウトを使用します。
*   画面幅がやや広い場合に表示される `TwoColumns`。 Web 上の Outlook は、ほとんどのタブレットでこのビューを使用します。
*   画面幅が広い場合に表示される `ThreeColumns`。 たとえば、Outlook on the web では、このビューをデスクトップコンピューターの全画面表示ウィンドウで使用します。

##### <a name="type"></a>型

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|
