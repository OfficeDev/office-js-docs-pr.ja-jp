---
title: Office.--の要件セット1.5
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 9ecbf4382f10b86ecdea41706211094029be09d2
ms.sourcegitcommit: dc78ee2a89fe3d4cd6f748be1eec9081c1077502
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2019
ms.locfileid: "36231257"
---
# <a name="diagnostics"></a><span data-ttu-id="546ba-102">診断</span><span class="sxs-lookup"><span data-stu-id="546ba-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="546ba-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="546ba-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="546ba-104">Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="546ba-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="546ba-105">要件</span><span class="sxs-lookup"><span data-stu-id="546ba-105">Requirements</span></span>

|<span data-ttu-id="546ba-106">要件</span><span class="sxs-lookup"><span data-stu-id="546ba-106">Requirement</span></span>| <span data-ttu-id="546ba-107">値</span><span class="sxs-lookup"><span data-stu-id="546ba-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="546ba-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="546ba-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="546ba-109">1.0</span><span class="sxs-lookup"><span data-stu-id="546ba-109">1.0</span></span>|
|[<span data-ttu-id="546ba-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="546ba-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="546ba-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="546ba-111">ReadItem</span></span>|
|[<span data-ttu-id="546ba-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="546ba-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="546ba-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="546ba-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="546ba-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="546ba-114">Members and methods</span></span>

| <span data-ttu-id="546ba-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="546ba-115">Member</span></span> | <span data-ttu-id="546ba-116">種類</span><span class="sxs-lookup"><span data-stu-id="546ba-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="546ba-117">名</span><span class="sxs-lookup"><span data-stu-id="546ba-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="546ba-118">Member</span><span class="sxs-lookup"><span data-stu-id="546ba-118">Member</span></span> |
| [<span data-ttu-id="546ba-119">上 diagnostics.hostversion</span><span class="sxs-lookup"><span data-stu-id="546ba-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="546ba-120">Member</span><span class="sxs-lookup"><span data-stu-id="546ba-120">Member</span></span> |
| [<span data-ttu-id="546ba-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="546ba-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="546ba-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="546ba-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="546ba-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="546ba-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="546ba-124">hostName: String</span><span class="sxs-lookup"><span data-stu-id="546ba-124">hostName: String</span></span>

<span data-ttu-id="546ba-125">ホスト アプリケーションの名前を表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="546ba-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="546ba-126">文字列は、値 `Outlook`、`OutlookIOS`、`OutlookWebApp` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="546ba-126">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="546ba-127">型</span><span class="sxs-lookup"><span data-stu-id="546ba-127">Type</span></span>

*   <span data-ttu-id="546ba-128">String</span><span class="sxs-lookup"><span data-stu-id="546ba-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="546ba-129">要件</span><span class="sxs-lookup"><span data-stu-id="546ba-129">Requirements</span></span>

|<span data-ttu-id="546ba-130">要件</span><span class="sxs-lookup"><span data-stu-id="546ba-130">Requirement</span></span>| <span data-ttu-id="546ba-131">値</span><span class="sxs-lookup"><span data-stu-id="546ba-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="546ba-132">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="546ba-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="546ba-133">1.0</span><span class="sxs-lookup"><span data-stu-id="546ba-133">1.0</span></span>|
|[<span data-ttu-id="546ba-134">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="546ba-134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="546ba-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="546ba-135">ReadItem</span></span>|
|[<span data-ttu-id="546ba-136">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="546ba-136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="546ba-137">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="546ba-137">Compose or Read</span></span>|

#### <a name="hostversion-string"></a><span data-ttu-id="546ba-138">hostVersion: String</span><span class="sxs-lookup"><span data-stu-id="546ba-138">hostVersion: String</span></span>

<span data-ttu-id="546ba-139">ホスト アプリケーションまたは Exchange Server のバージョンを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="546ba-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="546ba-140">メールアドインが Outlook デスクトップクライアントまたは iOS で実行されている場合、 `hostVersion`このプロパティはホストアプリケーションのバージョン (outlook) を返します。</span><span class="sxs-lookup"><span data-stu-id="546ba-140">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="546ba-141">Web 上の Outlook では、このプロパティは Exchange サーバーのバージョンを返します。</span><span class="sxs-lookup"><span data-stu-id="546ba-141">In Outlook on the web, the property returns the version of the Exchange Server.</span></span> <span data-ttu-id="546ba-142">例として、"15.0.468.0" という文字列があります。</span><span class="sxs-lookup"><span data-stu-id="546ba-142">An example is the string "15.0.468.0".</span></span>

##### <a name="type"></a><span data-ttu-id="546ba-143">型</span><span class="sxs-lookup"><span data-stu-id="546ba-143">Type</span></span>

*   <span data-ttu-id="546ba-144">String</span><span class="sxs-lookup"><span data-stu-id="546ba-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="546ba-145">要件</span><span class="sxs-lookup"><span data-stu-id="546ba-145">Requirements</span></span>

|<span data-ttu-id="546ba-146">要件</span><span class="sxs-lookup"><span data-stu-id="546ba-146">Requirement</span></span>| <span data-ttu-id="546ba-147">値</span><span class="sxs-lookup"><span data-stu-id="546ba-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="546ba-148">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="546ba-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="546ba-149">1.0</span><span class="sxs-lookup"><span data-stu-id="546ba-149">1.0</span></span>|
|[<span data-ttu-id="546ba-150">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="546ba-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="546ba-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="546ba-151">ReadItem</span></span>|
|[<span data-ttu-id="546ba-152">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="546ba-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="546ba-153">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="546ba-153">Compose or Read</span></span>|

#### <a name="owaview-string"></a><span data-ttu-id="546ba-154">OWAView: String</span><span class="sxs-lookup"><span data-stu-id="546ba-154">OWAView: String</span></span>

<span data-ttu-id="546ba-155">Web 上の Outlook の現在のビューを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="546ba-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="546ba-156">返される文字列は、値 `OneColumn`、`TwoColumns`、または `ThreeColumns` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="546ba-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="546ba-157">ホストアプリケーションが web 上の Outlook ではない場合、このプロパティにアクセスする`undefined`と、になります。</span><span class="sxs-lookup"><span data-stu-id="546ba-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="546ba-158">Outlook on the web には、画面とウィンドウの幅、および表示できる列の数に対応する3つのビューがあります。</span><span class="sxs-lookup"><span data-stu-id="546ba-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="546ba-159">画面幅が狭い場合に表示される `OneColumn`。</span><span class="sxs-lookup"><span data-stu-id="546ba-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="546ba-160">Outlook on the web では、スマートフォンの画面全体でこのような単一の列のレイアウトを使用します。</span><span class="sxs-lookup"><span data-stu-id="546ba-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="546ba-161">画面幅がやや広い場合に表示される `TwoColumns`。</span><span class="sxs-lookup"><span data-stu-id="546ba-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="546ba-162">Web 上の Outlook は、ほとんどのタブレットでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="546ba-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="546ba-163">画面幅が広い場合に表示される `ThreeColumns`。</span><span class="sxs-lookup"><span data-stu-id="546ba-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="546ba-164">たとえば、Outlook on the web では、このビューをデスクトップコンピューターの全画面表示ウィンドウで使用します。</span><span class="sxs-lookup"><span data-stu-id="546ba-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="546ba-165">型</span><span class="sxs-lookup"><span data-stu-id="546ba-165">Type</span></span>

*   <span data-ttu-id="546ba-166">String</span><span class="sxs-lookup"><span data-stu-id="546ba-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="546ba-167">要件</span><span class="sxs-lookup"><span data-stu-id="546ba-167">Requirements</span></span>

|<span data-ttu-id="546ba-168">要件</span><span class="sxs-lookup"><span data-stu-id="546ba-168">Requirement</span></span>| <span data-ttu-id="546ba-169">値</span><span class="sxs-lookup"><span data-stu-id="546ba-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="546ba-170">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="546ba-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="546ba-171">1.0</span><span class="sxs-lookup"><span data-stu-id="546ba-171">1.0</span></span>|
|[<span data-ttu-id="546ba-172">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="546ba-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="546ba-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="546ba-173">ReadItem</span></span>|
|[<span data-ttu-id="546ba-174">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="546ba-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="546ba-175">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="546ba-175">Compose or Read</span></span>|
