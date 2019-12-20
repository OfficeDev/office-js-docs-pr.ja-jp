---
title: Outlook アドイン API 要件セットのプレビュー
description: ''
ms.date: 12/17/2019
localization_priority: Priority
ms.openlocfilehash: a3cc49562add2f6fe54cf83d2f2ed64ebb61d8c7
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40815047"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook アドイン API 要件セットのプレビュー

JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。

> [!IMPORTANT]
> このドキュメントは、[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)の**プレビュー**用です。 この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。 アドイン マニフェストでこの要件を指定しないでください。

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

要件セットのプレビューには、[要件セット 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) のすべての機能が含まれています。

## <a name="features-in-preview"></a>プレビューの機能

次の機能はプレビュー段階です。

### <a name="integration-with-actionable-messages"></a>操作可能なメッセージとの統合

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdmethods"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (クラシック)

<br>

---

---

### <a name="office-theme"></a>Office テーマ

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Office テーマを取得する機能が追加されました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

`OfficeThemeChanged` イベントが `Mailbox` に追加されました。

**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)

<br>

---

### <a name="sso"></a>SSO

#### <a name="officeruntimeauthgetaccesstokenofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[OfficeRuntime.auth.getAccessToken](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

Microsoft Graph API の[アクセス トークンの取得](/outlook/add-ins/authenticate-a-user-with-an-sso-token)をアドインに対して許可する、`getAccessToken` へのアクセスが追加されました。

**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac、Outlook on the web (モダン)、Outlook on the web (クラシック)

## <a name="see-also"></a>関連項目

- [Outlook アドイン](/outlook/add-ins/)
- [Outlook アドインのコード サンプル](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [概要](/outlook/add-ins/quick-start)
- [要求セットとサポートされているクライアント](../../requirement-sets/outlook-api-requirement-sets.md)
