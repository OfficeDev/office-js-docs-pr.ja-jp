---
title: Office.--の要件セット1.6
description: ''
ms.date: 08/05/2019
localization_priority: Normal
ms.openlocfilehash: 63a9e453d448899c2dd3e12ca6cd0a83c47e66e5
ms.sourcegitcommit: dc78ee2a89fe3d4cd6f748be1eec9081c1077502
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2019
ms.locfileid: "36231250"
---
# <a name="diagnostics"></a><span data-ttu-id="84f96-102">診断</span><span class="sxs-lookup"><span data-stu-id="84f96-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="84f96-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="84f96-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="84f96-104">Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="84f96-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="84f96-105">要件</span><span class="sxs-lookup"><span data-stu-id="84f96-105">Requirements</span></span>

|<span data-ttu-id="84f96-106">要件</span><span class="sxs-lookup"><span data-stu-id="84f96-106">Requirement</span></span>| <span data-ttu-id="84f96-107">値</span><span class="sxs-lookup"><span data-stu-id="84f96-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="84f96-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="84f96-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="84f96-109">1.0</span><span class="sxs-lookup"><span data-stu-id="84f96-109">1.0</span></span>|
|[<span data-ttu-id="84f96-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="84f96-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="84f96-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="84f96-111">ReadItem</span></span>|
|[<span data-ttu-id="84f96-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="84f96-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="84f96-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="84f96-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="84f96-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="84f96-114">Members and methods</span></span>

| <span data-ttu-id="84f96-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="84f96-115">Member</span></span> | <span data-ttu-id="84f96-116">種類</span><span class="sxs-lookup"><span data-stu-id="84f96-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="84f96-117">名</span><span class="sxs-lookup"><span data-stu-id="84f96-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="84f96-118">Member</span><span class="sxs-lookup"><span data-stu-id="84f96-118">Member</span></span> |
| [<span data-ttu-id="84f96-119">上 diagnostics.hostversion</span><span class="sxs-lookup"><span data-stu-id="84f96-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="84f96-120">Member</span><span class="sxs-lookup"><span data-stu-id="84f96-120">Member</span></span> |
| [<span data-ttu-id="84f96-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="84f96-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="84f96-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="84f96-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="84f96-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="84f96-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="84f96-124">hostName: String</span><span class="sxs-lookup"><span data-stu-id="84f96-124">hostName: String</span></span>

<span data-ttu-id="84f96-125">ホスト アプリケーションの名前を表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="84f96-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="84f96-126">文字列は、値 `Outlook`、`OutlookWebApp`、`OutlookIOS`、または `OutlookAndroid` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="84f96-126">A string that can be one of the following values: `Outlook`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.</span></span>

##### <a name="type"></a><span data-ttu-id="84f96-127">型</span><span class="sxs-lookup"><span data-stu-id="84f96-127">Type</span></span>

*   <span data-ttu-id="84f96-128">String</span><span class="sxs-lookup"><span data-stu-id="84f96-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="84f96-129">要件</span><span class="sxs-lookup"><span data-stu-id="84f96-129">Requirements</span></span>

|<span data-ttu-id="84f96-130">要件</span><span class="sxs-lookup"><span data-stu-id="84f96-130">Requirement</span></span>| <span data-ttu-id="84f96-131">値</span><span class="sxs-lookup"><span data-stu-id="84f96-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="84f96-132">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="84f96-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="84f96-133">1.0</span><span class="sxs-lookup"><span data-stu-id="84f96-133">1.0</span></span>|
|[<span data-ttu-id="84f96-134">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="84f96-134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="84f96-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="84f96-135">ReadItem</span></span>|
|[<span data-ttu-id="84f96-136">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="84f96-136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="84f96-137">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="84f96-137">Compose or Read</span></span>|

#### <a name="hostversion-string"></a><span data-ttu-id="84f96-138">hostVersion: String</span><span class="sxs-lookup"><span data-stu-id="84f96-138">hostVersion: String</span></span>

<span data-ttu-id="84f96-139">ホスト アプリケーションまたは Exchange Server のバージョンを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="84f96-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="84f96-140">メールアドインが Outlook デスクトップクライアントまたは iOS で実行されている場合、 `hostVersion`このプロパティはホストアプリケーションのバージョン (outlook) を返します。</span><span class="sxs-lookup"><span data-stu-id="84f96-140">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="84f96-141">Web 上の Outlook では、このプロパティは Exchange サーバーのバージョンを返します。</span><span class="sxs-lookup"><span data-stu-id="84f96-141">In Outlook on the web, the property returns the version of the Exchange Server.</span></span> <span data-ttu-id="84f96-142">例として、"15.0.468.0" という文字列があります。</span><span class="sxs-lookup"><span data-stu-id="84f96-142">An example is the string "15.0.468.0".</span></span>

##### <a name="type"></a><span data-ttu-id="84f96-143">型</span><span class="sxs-lookup"><span data-stu-id="84f96-143">Type</span></span>

*   <span data-ttu-id="84f96-144">String</span><span class="sxs-lookup"><span data-stu-id="84f96-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="84f96-145">要件</span><span class="sxs-lookup"><span data-stu-id="84f96-145">Requirements</span></span>

|<span data-ttu-id="84f96-146">要件</span><span class="sxs-lookup"><span data-stu-id="84f96-146">Requirement</span></span>| <span data-ttu-id="84f96-147">値</span><span class="sxs-lookup"><span data-stu-id="84f96-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="84f96-148">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="84f96-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="84f96-149">1.0</span><span class="sxs-lookup"><span data-stu-id="84f96-149">1.0</span></span>|
|[<span data-ttu-id="84f96-150">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="84f96-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="84f96-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="84f96-151">ReadItem</span></span>|
|[<span data-ttu-id="84f96-152">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="84f96-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="84f96-153">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="84f96-153">Compose or Read</span></span>|

#### <a name="owaview-string"></a><span data-ttu-id="84f96-154">OWAView: String</span><span class="sxs-lookup"><span data-stu-id="84f96-154">OWAView: String</span></span>

<span data-ttu-id="84f96-155">Web 上の Outlook の現在のビューを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="84f96-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="84f96-156">返される文字列は、値 `OneColumn`、`TwoColumns`、または `ThreeColumns` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="84f96-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="84f96-157">ホストアプリケーションが web 上の Outlook ではない場合、このプロパティにアクセスする`undefined`と、になります。</span><span class="sxs-lookup"><span data-stu-id="84f96-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="84f96-158">Outlook on the web には、画面とウィンドウの幅、および表示できる列の数に対応する3つのビューがあります。</span><span class="sxs-lookup"><span data-stu-id="84f96-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="84f96-159">画面幅が狭い場合に表示される `OneColumn`。</span><span class="sxs-lookup"><span data-stu-id="84f96-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="84f96-160">Outlook on the web では、スマートフォンの画面全体でこのような単一の列のレイアウトを使用します。</span><span class="sxs-lookup"><span data-stu-id="84f96-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="84f96-161">画面幅がやや広い場合に表示される `TwoColumns`。</span><span class="sxs-lookup"><span data-stu-id="84f96-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="84f96-162">Web 上の Outlook は、ほとんどのタブレットでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="84f96-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="84f96-163">画面幅が広い場合に表示される `ThreeColumns`。</span><span class="sxs-lookup"><span data-stu-id="84f96-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="84f96-164">たとえば、Outlook on the web では、このビューをデスクトップコンピューターの全画面表示ウィンドウで使用します。</span><span class="sxs-lookup"><span data-stu-id="84f96-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="84f96-165">型</span><span class="sxs-lookup"><span data-stu-id="84f96-165">Type</span></span>

*   <span data-ttu-id="84f96-166">String</span><span class="sxs-lookup"><span data-stu-id="84f96-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="84f96-167">要件</span><span class="sxs-lookup"><span data-stu-id="84f96-167">Requirements</span></span>

|<span data-ttu-id="84f96-168">要件</span><span class="sxs-lookup"><span data-stu-id="84f96-168">Requirement</span></span>| <span data-ttu-id="84f96-169">値</span><span class="sxs-lookup"><span data-stu-id="84f96-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="84f96-170">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="84f96-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="84f96-171">1.0</span><span class="sxs-lookup"><span data-stu-id="84f96-171">1.0</span></span>|
|[<span data-ttu-id="84f96-172">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="84f96-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="84f96-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="84f96-173">ReadItem</span></span>|
|[<span data-ttu-id="84f96-174">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="84f96-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="84f96-175">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="84f96-175">Compose or Read</span></span>|
