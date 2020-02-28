---
title: Outlook アドイン API 要件セットのプレビュー
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 87c15ac889a955412e6a8350baaed8611fdb5164
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325221"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="046d2-102">Outlook アドイン API 要件セットのプレビュー</span><span class="sxs-lookup"><span data-stu-id="046d2-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="046d2-103">Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="046d2-103">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="046d2-104">このドキュメントは、[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)の**プレビュー**用です。</span><span class="sxs-lookup"><span data-stu-id="046d2-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="046d2-105">この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。</span><span class="sxs-lookup"><span data-stu-id="046d2-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="046d2-106">アドイン マニフェストでこの要件を指定しないでください。</span><span class="sxs-lookup"><span data-stu-id="046d2-106">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="046d2-107">要件セットのプレビューには、[要件セット 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) のすべての機能が含まれています。</span><span class="sxs-lookup"><span data-stu-id="046d2-107">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="046d2-108">プレビューの機能</span><span class="sxs-lookup"><span data-stu-id="046d2-108">Features in preview</span></span>

<span data-ttu-id="046d2-109">次の機能はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="046d2-109">The following features are in preview.</span></span>

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="046d2-110">操作可能なメッセージとの統合</span><span class="sxs-lookup"><span data-stu-id="046d2-110">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="046d2-111">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="046d2-111">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="046d2-112">アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。</span><span class="sxs-lookup"><span data-stu-id="046d2-112">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="046d2-113">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="046d2-113">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="046d2-114">Office テーマ</span><span class="sxs-lookup"><span data-stu-id="046d2-114">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="046d2-115">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="046d2-115">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="046d2-116">Office テーマを取得する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="046d2-116">Added ability to get Office theme.</span></span>

<span data-ttu-id="046d2-117">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="046d2-117">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="046d2-118">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="046d2-118">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="046d2-119">`OfficeThemeChanged` イベントが `Mailbox` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="046d2-119">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="046d2-120">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="046d2-120">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="sso"></a><span data-ttu-id="046d2-121">SSO</span><span class="sxs-lookup"><span data-stu-id="046d2-121">SSO</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="046d2-122">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="046d2-122">OfficeRuntime.auth.getAccessToken</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="046d2-123">Microsoft Graph API の[アクセス トークンの取得](../../../outlook/authenticate-a-user-with-an-sso-token.md)をアドインに対して許可する、`getAccessToken` へのアクセスが追加されました。</span><span class="sxs-lookup"><span data-stu-id="046d2-123">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="046d2-124">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac、Outlook on the web (モダン)、Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="046d2-124">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="046d2-125">関連項目</span><span class="sxs-lookup"><span data-stu-id="046d2-125">See also</span></span>

- [<span data-ttu-id="046d2-126">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="046d2-126">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="046d2-127">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="046d2-127">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="046d2-128">概要</span><span class="sxs-lookup"><span data-stu-id="046d2-128">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="046d2-129">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="046d2-129">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
