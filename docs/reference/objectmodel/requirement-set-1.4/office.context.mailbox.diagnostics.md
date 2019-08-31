---
title: Office.--の要件セット1.4
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 9c6543cad17d464ce139381270ad1495d43d5cd9
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696261"
---
# <a name="diagnostics"></a><span data-ttu-id="ef515-102">診断</span><span class="sxs-lookup"><span data-stu-id="ef515-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="ef515-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="ef515-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="ef515-104">Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="ef515-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ef515-105">要件</span><span class="sxs-lookup"><span data-stu-id="ef515-105">Requirements</span></span>

|<span data-ttu-id="ef515-106">要件</span><span class="sxs-lookup"><span data-stu-id="ef515-106">Requirement</span></span>| <span data-ttu-id="ef515-107">値</span><span class="sxs-lookup"><span data-stu-id="ef515-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ef515-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ef515-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ef515-109">1.0</span><span class="sxs-lookup"><span data-stu-id="ef515-109">1.0</span></span>|
|[<span data-ttu-id="ef515-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ef515-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ef515-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ef515-111">ReadItem</span></span>|
|[<span data-ttu-id="ef515-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ef515-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ef515-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ef515-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ef515-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="ef515-114">Members and methods</span></span>

| <span data-ttu-id="ef515-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="ef515-115">Member</span></span> | <span data-ttu-id="ef515-116">種類</span><span class="sxs-lookup"><span data-stu-id="ef515-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ef515-117">名</span><span class="sxs-lookup"><span data-stu-id="ef515-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="ef515-118">Member</span><span class="sxs-lookup"><span data-stu-id="ef515-118">Member</span></span> |
| [<span data-ttu-id="ef515-119">上 diagnostics.hostversion</span><span class="sxs-lookup"><span data-stu-id="ef515-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="ef515-120">Member</span><span class="sxs-lookup"><span data-stu-id="ef515-120">Member</span></span> |
| [<span data-ttu-id="ef515-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="ef515-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="ef515-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="ef515-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="ef515-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="ef515-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="ef515-124">hostName: String</span><span class="sxs-lookup"><span data-stu-id="ef515-124">hostName: String</span></span>

<span data-ttu-id="ef515-125">ホスト アプリケーションの名前を表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="ef515-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="ef515-126">文字列は、値 `Outlook`、`OutlookIOS`、`OutlookWebApp` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="ef515-126">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

> [!NOTE]
> <span data-ttu-id="ef515-127">この`Outlook`値は、デスクトップクライアント (つまり Windows と Mac) の Outlook に対して返されます。</span><span class="sxs-lookup"><span data-stu-id="ef515-127">The `Outlook` value is returned for Outlook on desktop clients (i.e., Windows and Mac).</span></span>

##### <a name="type"></a><span data-ttu-id="ef515-128">型</span><span class="sxs-lookup"><span data-stu-id="ef515-128">Type</span></span>

*   <span data-ttu-id="ef515-129">String</span><span class="sxs-lookup"><span data-stu-id="ef515-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ef515-130">要件</span><span class="sxs-lookup"><span data-stu-id="ef515-130">Requirements</span></span>

|<span data-ttu-id="ef515-131">要件</span><span class="sxs-lookup"><span data-stu-id="ef515-131">Requirement</span></span>| <span data-ttu-id="ef515-132">値</span><span class="sxs-lookup"><span data-stu-id="ef515-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="ef515-133">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ef515-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ef515-134">1.0</span><span class="sxs-lookup"><span data-stu-id="ef515-134">1.0</span></span>|
|[<span data-ttu-id="ef515-135">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ef515-135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ef515-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ef515-136">ReadItem</span></span>|
|[<span data-ttu-id="ef515-137">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ef515-137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ef515-138">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ef515-138">Compose or Read</span></span>|

<br>

---
---

#### <a name="hostversion-string"></a><span data-ttu-id="ef515-139">hostVersion: String</span><span class="sxs-lookup"><span data-stu-id="ef515-139">hostVersion: String</span></span>

<span data-ttu-id="ef515-140">ホストアプリケーションまたは Exchange サーバー (例: "15.0.468.0") のいずれかのバージョンを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="ef515-140">Gets a string that represents the version of either the host application or the Exchange Server (e.g. "15.0.468.0").</span></span>

<span data-ttu-id="ef515-141">メールアドインが Outlook デスクトップクライアントまたは iOS で実行されている場合、 `hostVersion`このプロパティはホストアプリケーションのバージョン (outlook) を返します。</span><span class="sxs-lookup"><span data-stu-id="ef515-141">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="ef515-142">Web 上の Outlook では、このプロパティは Exchange サーバーのバージョンを返します。</span><span class="sxs-lookup"><span data-stu-id="ef515-142">In Outlook on the web, the property returns the version of the Exchange Server.</span></span>

##### <a name="type"></a><span data-ttu-id="ef515-143">型</span><span class="sxs-lookup"><span data-stu-id="ef515-143">Type</span></span>

*   <span data-ttu-id="ef515-144">String</span><span class="sxs-lookup"><span data-stu-id="ef515-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ef515-145">要件</span><span class="sxs-lookup"><span data-stu-id="ef515-145">Requirements</span></span>

|<span data-ttu-id="ef515-146">要件</span><span class="sxs-lookup"><span data-stu-id="ef515-146">Requirement</span></span>| <span data-ttu-id="ef515-147">値</span><span class="sxs-lookup"><span data-stu-id="ef515-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="ef515-148">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ef515-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ef515-149">1.0</span><span class="sxs-lookup"><span data-stu-id="ef515-149">1.0</span></span>|
|[<span data-ttu-id="ef515-150">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ef515-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ef515-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ef515-151">ReadItem</span></span>|
|[<span data-ttu-id="ef515-152">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ef515-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ef515-153">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ef515-153">Compose or Read</span></span>|

<br>

---
---

#### <a name="owaview-string"></a><span data-ttu-id="ef515-154">OWAView: String</span><span class="sxs-lookup"><span data-stu-id="ef515-154">OWAView: String</span></span>

<span data-ttu-id="ef515-155">Web 上の Outlook の現在のビューを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="ef515-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="ef515-156">返される文字列は、値 `OneColumn`、`TwoColumns`、または `ThreeColumns` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="ef515-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="ef515-157">ホストアプリケーションが web 上の Outlook ではない場合、このプロパティにアクセスする`undefined`と、になります。</span><span class="sxs-lookup"><span data-stu-id="ef515-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="ef515-158">Outlook on the web には、画面とウィンドウの幅、および表示できる列の数に対応する3つのビューがあります。</span><span class="sxs-lookup"><span data-stu-id="ef515-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="ef515-159">画面幅が狭い場合に表示される `OneColumn`。</span><span class="sxs-lookup"><span data-stu-id="ef515-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="ef515-160">Outlook on the web では、スマートフォンの画面全体でこのような単一の列のレイアウトを使用します。</span><span class="sxs-lookup"><span data-stu-id="ef515-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="ef515-161">画面幅がやや広い場合に表示される `TwoColumns`。</span><span class="sxs-lookup"><span data-stu-id="ef515-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="ef515-162">Web 上の Outlook は、ほとんどのタブレットでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="ef515-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="ef515-163">画面幅が広い場合に表示される `ThreeColumns`。</span><span class="sxs-lookup"><span data-stu-id="ef515-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="ef515-164">たとえば、Outlook on the web では、このビューをデスクトップコンピューターの全画面表示ウィンドウで使用します。</span><span class="sxs-lookup"><span data-stu-id="ef515-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="ef515-165">型</span><span class="sxs-lookup"><span data-stu-id="ef515-165">Type</span></span>

*   <span data-ttu-id="ef515-166">String</span><span class="sxs-lookup"><span data-stu-id="ef515-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ef515-167">要件</span><span class="sxs-lookup"><span data-stu-id="ef515-167">Requirements</span></span>

|<span data-ttu-id="ef515-168">要件</span><span class="sxs-lookup"><span data-stu-id="ef515-168">Requirement</span></span>| <span data-ttu-id="ef515-169">値</span><span class="sxs-lookup"><span data-stu-id="ef515-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="ef515-170">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ef515-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ef515-171">1.0</span><span class="sxs-lookup"><span data-stu-id="ef515-171">1.0</span></span>|
|[<span data-ttu-id="ef515-172">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ef515-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ef515-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ef515-173">ReadItem</span></span>|
|[<span data-ttu-id="ef515-174">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ef515-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ef515-175">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ef515-175">Compose or Read</span></span>|
