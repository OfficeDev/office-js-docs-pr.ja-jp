---
title: Office の設定-プレビュー要件セット
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 492e292737417854adfaf98feb2b67788933d874
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629203"
---
# <a name="diagnostics"></a>診断

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a>[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics

Outlook アドインに診断情報を提供します。

##### <a name="requirements"></a>Requirements

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

##### <a name="properties"></a>プロパティ

| プロパティ | 最小値<br>アクセス許可レベル | モード | 戻り値の種類 | 最小値<br>要件セット |
|---|---|---|---|---|
| [名](#hostname-string) | ReadItem | 作成<br>読み取り | String | 1.0 |
| [上 diagnostics.hostversion](#hostversion-string) | ReadItem | 作成<br>読み取り | String | 1.0 |
| [OWAView](#owaview-string) | ReadItem | 作成<br>読み取り | String | 1.0 |

## <a name="property-details"></a>プロパティの詳細

#### <a name="hostname-string"></a>hostName: String

ホスト アプリケーションの名前を表す文字列を取得します。

文字列は、値 `Outlook`、`OutlookWebApp`、`OutlookIOS`、または `OutlookAndroid` のいずれかになります。

> [!NOTE]
> この`Outlook`値は、デスクトップクライアント (つまり Windows と Mac) の Outlook に対して返されます。

##### <a name="type"></a>型

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

<br>

---
---

#### <a name="hostversion-string"></a>hostVersion: String

ホストアプリケーションまたは Exchange サーバー (例: "15.0.468.0") のバージョンを表す文字列を取得します。

メールアドインが Outlook デスクトップまたはモバイルクライアント上で実行されている場合`hostVersion` 、このプロパティはホストアプリケーションのバージョン (outlook) を返します。 Web 上の Outlook では、このプロパティは Exchange サーバーのバージョンを返します。

##### <a name="type"></a>型

*   String

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](/outlook/add-ins/#extension-points)| 新規作成または閲覧|

<br>

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
