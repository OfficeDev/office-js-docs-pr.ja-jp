---
title: Outlook アドイン API 要件セットのプレビュー
description: ''
ms.date: 03/04/2020
localization_priority: Normal
ms.openlocfilehash: 4365dab3d8dd1ddb876536b3030926d68a89ac49
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/12/2020
ms.locfileid: "42605674"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="2b056-102">Outlook アドイン API 要件セットのプレビュー</span><span class="sxs-lookup"><span data-stu-id="2b056-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="2b056-103">Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="2b056-103">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2b056-104">このドキュメントは、[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)の**プレビュー**用です。</span><span class="sxs-lookup"><span data-stu-id="2b056-104">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="2b056-105">この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。</span><span class="sxs-lookup"><span data-stu-id="2b056-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="2b056-106">アドイン マニフェストでこの要件を指定しないでください。</span><span class="sxs-lookup"><span data-stu-id="2b056-106">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="2b056-107">要件セットのプレビューには、[要件セット 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) のすべての機能が含まれています。</span><span class="sxs-lookup"><span data-stu-id="2b056-107">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="2b056-108">プレビューの機能</span><span class="sxs-lookup"><span data-stu-id="2b056-108">Features in preview</span></span>

<span data-ttu-id="2b056-109">次の機能はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="2b056-109">The following features are in preview.</span></span>

### <a name="append-on-send"></a><span data-ttu-id="2b056-110">送信時に追加</span><span class="sxs-lookup"><span data-stu-id="2b056-110">Append on send</span></span>

#### <a name="officebodyappendonsendasync"></a>[<span data-ttu-id="2b056-111">Office. appendOnSendAsync</span><span class="sxs-lookup"><span data-stu-id="2b056-111">Office.Body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="2b056-112">新規作成モードで、アイテム`Body`の本文の最後にデータを追加する新しい関数をオブジェクトに追加しました。</span><span class="sxs-lookup"><span data-stu-id="2b056-112">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="2b056-113">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="2b056-113">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="2b056-114">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="2b056-114">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="2b056-115">拡張されたアクセス許可のコレクションに`AppendOnSend`拡張アクセス許可が含まれている必要があるマニフェストに、新しい要素を追加しました。</span><span class="sxs-lookup"><span data-stu-id="2b056-115">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="2b056-116">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="2b056-116">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="2b056-117">操作可能なメッセージとの統合</span><span class="sxs-lookup"><span data-stu-id="2b056-117">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="2b056-118">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="2b056-118">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="2b056-119">アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。</span><span class="sxs-lookup"><span data-stu-id="2b056-119">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="2b056-120">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="2b056-120">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="2b056-121">Office テーマ</span><span class="sxs-lookup"><span data-stu-id="2b056-121">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="2b056-122">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="2b056-122">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="2b056-123">Office テーマを取得する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="2b056-123">Added ability to get Office theme.</span></span>

<span data-ttu-id="2b056-124">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="2b056-124">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="2b056-125">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="2b056-125">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="2b056-126">`OfficeThemeChanged` イベントが `Mailbox` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="2b056-126">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="2b056-127">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="2b056-127">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="sso"></a><span data-ttu-id="2b056-128">SSO</span><span class="sxs-lookup"><span data-stu-id="2b056-128">SSO</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="2b056-129">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="2b056-129">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="2b056-130">Microsoft Graph API の[アクセス トークンの取得](../../../outlook/authenticate-a-user-with-an-sso-token.md)をアドインに対して許可する、`getAccessToken` へのアクセスが追加されました。</span><span class="sxs-lookup"><span data-stu-id="2b056-130">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="2b056-131">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac、Outlook on the web (モダン)、Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="2b056-131">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="2b056-132">関連項目</span><span class="sxs-lookup"><span data-stu-id="2b056-132">See also</span></span>

- [<span data-ttu-id="2b056-133">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="2b056-133">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="2b056-134">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="2b056-134">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="2b056-135">概要</span><span class="sxs-lookup"><span data-stu-id="2b056-135">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="2b056-136">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="2b056-136">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
