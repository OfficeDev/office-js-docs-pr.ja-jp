---
title: Outlook アドイン API 要件セットのプレビュー
description: Outlook アドインおよび Office JavaScript Api で現在プレビューされている機能と Api。
ms.date: 03/04/2020
localization_priority: Normal
ms.openlocfilehash: c87ce8472becc072702f58e7d8c21665904673d2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717811"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="fd038-103">Outlook アドイン API 要件セットのプレビュー</span><span class="sxs-lookup"><span data-stu-id="fd038-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="fd038-104">Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="fd038-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fd038-105">このドキュメントは、[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)の**プレビュー**用です。</span><span class="sxs-lookup"><span data-stu-id="fd038-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="fd038-106">この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。</span><span class="sxs-lookup"><span data-stu-id="fd038-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="fd038-107">アドイン マニフェストでこの要件を指定しないでください。</span><span class="sxs-lookup"><span data-stu-id="fd038-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="fd038-108">要件セットのプレビューには、[要件セット 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) のすべての機能が含まれています。</span><span class="sxs-lookup"><span data-stu-id="fd038-108">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="fd038-109">プレビューの機能</span><span class="sxs-lookup"><span data-stu-id="fd038-109">Features in preview</span></span>

<span data-ttu-id="fd038-110">次の機能はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="fd038-110">The following features are in preview.</span></span>

### <a name="append-on-send"></a><span data-ttu-id="fd038-111">送信時に追加</span><span class="sxs-lookup"><span data-stu-id="fd038-111">Append on send</span></span>

#### <a name="officebodyappendonsendasync"></a>[<span data-ttu-id="fd038-112">Office. appendOnSendAsync</span><span class="sxs-lookup"><span data-stu-id="fd038-112">Office.Body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="fd038-113">新規作成モードで、アイテム`Body`の本文の最後にデータを追加する新しい関数をオブジェクトに追加しました。</span><span class="sxs-lookup"><span data-stu-id="fd038-113">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="fd038-114">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="fd038-114">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="fd038-115">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="fd038-115">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="fd038-116">拡張されたアクセス許可のコレクションに`AppendOnSend`拡張アクセス許可が含まれている必要があるマニフェストに、新しい要素を追加しました。</span><span class="sxs-lookup"><span data-stu-id="fd038-116">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="fd038-117">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="fd038-117">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="fd038-118">操作可能なメッセージとの統合</span><span class="sxs-lookup"><span data-stu-id="fd038-118">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="fd038-119">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="fd038-119">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="fd038-120">アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。</span><span class="sxs-lookup"><span data-stu-id="fd038-120">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="fd038-121">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="fd038-121">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="fd038-122">Office テーマ</span><span class="sxs-lookup"><span data-stu-id="fd038-122">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="fd038-123">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="fd038-123">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="fd038-124">Office テーマを取得する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="fd038-124">Added ability to get Office theme.</span></span>

<span data-ttu-id="fd038-125">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="fd038-125">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="fd038-126">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="fd038-126">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="fd038-127">`OfficeThemeChanged` イベントが `Mailbox` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="fd038-127">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="fd038-128">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="fd038-128">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="sso"></a><span data-ttu-id="fd038-129">SSO</span><span class="sxs-lookup"><span data-stu-id="fd038-129">SSO</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="fd038-130">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="fd038-130">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="fd038-131">Microsoft Graph API の[アクセス トークンの取得](../../../outlook/authenticate-a-user-with-an-sso-token.md)をアドインに対して許可する、`getAccessToken` へのアクセスが追加されました。</span><span class="sxs-lookup"><span data-stu-id="fd038-131">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="fd038-132">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac、Outlook on the web (モダン)、Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="fd038-132">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="fd038-133">関連項目</span><span class="sxs-lookup"><span data-stu-id="fd038-133">See also</span></span>

- [<span data-ttu-id="fd038-134">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="fd038-134">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="fd038-135">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="fd038-135">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="fd038-136">概要</span><span class="sxs-lookup"><span data-stu-id="fd038-136">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="fd038-137">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="fd038-137">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
