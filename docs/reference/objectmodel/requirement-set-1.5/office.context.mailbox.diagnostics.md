---
title: Office.--の要件セット1.5
description: ''
ms.date: 11/06/2019
localization_priority: Normal
ms.openlocfilehash: e69a4d7294cddd40bfc2c9393866561d2c1b9c62
ms.sourcegitcommit: 08c0b9ff319c391922fa43d3c2e9783cf6b53b1b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/08/2019
ms.locfileid: "38066138"
---
# <a name="diagnostics"></a><span data-ttu-id="479e1-102">診断</span><span class="sxs-lookup"><span data-stu-id="479e1-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="479e1-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="479e1-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="479e1-104">Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="479e1-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="479e1-105">要件</span><span class="sxs-lookup"><span data-stu-id="479e1-105">Requirements</span></span>

|<span data-ttu-id="479e1-106">要件</span><span class="sxs-lookup"><span data-stu-id="479e1-106">Requirement</span></span>| <span data-ttu-id="479e1-107">値</span><span class="sxs-lookup"><span data-stu-id="479e1-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="479e1-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="479e1-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="479e1-109">1.0</span><span class="sxs-lookup"><span data-stu-id="479e1-109">1.0</span></span>|
|[<span data-ttu-id="479e1-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="479e1-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="479e1-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="479e1-111">ReadItem</span></span>|
|[<span data-ttu-id="479e1-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="479e1-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="479e1-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="479e1-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="479e1-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="479e1-114">Members and methods</span></span>

| <span data-ttu-id="479e1-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="479e1-115">Member</span></span> | <span data-ttu-id="479e1-116">種類</span><span class="sxs-lookup"><span data-stu-id="479e1-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="479e1-117">名</span><span class="sxs-lookup"><span data-stu-id="479e1-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="479e1-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="479e1-118">Member</span></span> |
| [<span data-ttu-id="479e1-119">上 diagnostics.hostversion</span><span class="sxs-lookup"><span data-stu-id="479e1-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="479e1-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="479e1-120">Member</span></span> |
| [<span data-ttu-id="479e1-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="479e1-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="479e1-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="479e1-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="479e1-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="479e1-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="479e1-124">hostName: String</span><span class="sxs-lookup"><span data-stu-id="479e1-124">hostName: String</span></span>

<span data-ttu-id="479e1-125">ホスト アプリケーションの名前を表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="479e1-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="479e1-126">文字列は、値 `Outlook`、`OutlookWebApp`、`OutlookIOS`、または `OutlookAndroid` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="479e1-126">A string that can be one of the following values: `Outlook`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.</span></span>

> [!NOTE]
> <span data-ttu-id="479e1-127">この`Outlook`値は、デスクトップクライアント (つまり Windows と Mac) の Outlook に対して返されます。</span><span class="sxs-lookup"><span data-stu-id="479e1-127">The `Outlook` value is returned for Outlook on desktop clients (i.e., Windows and Mac).</span></span>

##### <a name="type"></a><span data-ttu-id="479e1-128">型</span><span class="sxs-lookup"><span data-stu-id="479e1-128">Type</span></span>

*   <span data-ttu-id="479e1-129">String</span><span class="sxs-lookup"><span data-stu-id="479e1-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="479e1-130">要件</span><span class="sxs-lookup"><span data-stu-id="479e1-130">Requirements</span></span>

|<span data-ttu-id="479e1-131">要件</span><span class="sxs-lookup"><span data-stu-id="479e1-131">Requirement</span></span>| <span data-ttu-id="479e1-132">値</span><span class="sxs-lookup"><span data-stu-id="479e1-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="479e1-133">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="479e1-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="479e1-134">1.0</span><span class="sxs-lookup"><span data-stu-id="479e1-134">1.0</span></span>|
|[<span data-ttu-id="479e1-135">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="479e1-135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="479e1-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="479e1-136">ReadItem</span></span>|
|[<span data-ttu-id="479e1-137">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="479e1-137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="479e1-138">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="479e1-138">Compose or Read</span></span>|

<br>

---
---

#### <a name="hostversion-string"></a><span data-ttu-id="479e1-139">hostVersion: String</span><span class="sxs-lookup"><span data-stu-id="479e1-139">hostVersion: String</span></span>

<span data-ttu-id="479e1-140">ホストアプリケーションまたは Exchange サーバー (例: "15.0.468.0") のバージョンを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="479e1-140">Gets a string that represents the version of either the host application or the Exchange Server (e.g., "15.0.468.0").</span></span>

<span data-ttu-id="479e1-141">メールアドインが Outlook デスクトップまたはモバイルクライアント上で実行されている場合`hostVersion` 、このプロパティはホストアプリケーションのバージョン (outlook) を返します。</span><span class="sxs-lookup"><span data-stu-id="479e1-141">If the mail add-in is running on an Outlook desktop or mobile client, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="479e1-142">Web 上の Outlook では、このプロパティは Exchange サーバーのバージョンを返します。</span><span class="sxs-lookup"><span data-stu-id="479e1-142">In Outlook on the web, the property returns the version of the Exchange Server.</span></span>

##### <a name="type"></a><span data-ttu-id="479e1-143">型</span><span class="sxs-lookup"><span data-stu-id="479e1-143">Type</span></span>

*   <span data-ttu-id="479e1-144">String</span><span class="sxs-lookup"><span data-stu-id="479e1-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="479e1-145">要件</span><span class="sxs-lookup"><span data-stu-id="479e1-145">Requirements</span></span>

|<span data-ttu-id="479e1-146">要件</span><span class="sxs-lookup"><span data-stu-id="479e1-146">Requirement</span></span>| <span data-ttu-id="479e1-147">値</span><span class="sxs-lookup"><span data-stu-id="479e1-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="479e1-148">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="479e1-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="479e1-149">1.0</span><span class="sxs-lookup"><span data-stu-id="479e1-149">1.0</span></span>|
|[<span data-ttu-id="479e1-150">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="479e1-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="479e1-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="479e1-151">ReadItem</span></span>|
|[<span data-ttu-id="479e1-152">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="479e1-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="479e1-153">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="479e1-153">Compose or Read</span></span>|

<br>

---
---

#### <a name="owaview-string"></a><span data-ttu-id="479e1-154">OWAView: String</span><span class="sxs-lookup"><span data-stu-id="479e1-154">OWAView: String</span></span>

<span data-ttu-id="479e1-155">Web 上の Outlook の現在のビューを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="479e1-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="479e1-156">返される文字列は、値 `OneColumn`、`TwoColumns`、または `ThreeColumns` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="479e1-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="479e1-157">ホストアプリケーションが web 上の Outlook ではない場合、このプロパティにアクセスする`undefined`と、になります。</span><span class="sxs-lookup"><span data-stu-id="479e1-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="479e1-158">Outlook on the web には、画面とウィンドウの幅、および表示できる列の数に対応する3つのビューがあります。</span><span class="sxs-lookup"><span data-stu-id="479e1-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="479e1-159">画面幅が狭い場合に表示される `OneColumn`。</span><span class="sxs-lookup"><span data-stu-id="479e1-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="479e1-160">Outlook on the web では、スマートフォンの画面全体でこのような単一の列のレイアウトを使用します。</span><span class="sxs-lookup"><span data-stu-id="479e1-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="479e1-161">画面幅がやや広い場合に表示される `TwoColumns`。</span><span class="sxs-lookup"><span data-stu-id="479e1-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="479e1-162">Web 上の Outlook は、ほとんどのタブレットでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="479e1-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="479e1-163">画面幅が広い場合に表示される `ThreeColumns`。</span><span class="sxs-lookup"><span data-stu-id="479e1-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="479e1-164">たとえば、Outlook on the web では、このビューをデスクトップコンピューターの全画面表示ウィンドウで使用します。</span><span class="sxs-lookup"><span data-stu-id="479e1-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="479e1-165">型</span><span class="sxs-lookup"><span data-stu-id="479e1-165">Type</span></span>

*   <span data-ttu-id="479e1-166">String</span><span class="sxs-lookup"><span data-stu-id="479e1-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="479e1-167">要件</span><span class="sxs-lookup"><span data-stu-id="479e1-167">Requirements</span></span>

|<span data-ttu-id="479e1-168">要件</span><span class="sxs-lookup"><span data-stu-id="479e1-168">Requirement</span></span>| <span data-ttu-id="479e1-169">値</span><span class="sxs-lookup"><span data-stu-id="479e1-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="479e1-170">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="479e1-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="479e1-171">1.0</span><span class="sxs-lookup"><span data-stu-id="479e1-171">1.0</span></span>|
|[<span data-ttu-id="479e1-172">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="479e1-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="479e1-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="479e1-173">ReadItem</span></span>|
|[<span data-ttu-id="479e1-174">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="479e1-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="479e1-175">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="479e1-175">Compose or Read</span></span>|
