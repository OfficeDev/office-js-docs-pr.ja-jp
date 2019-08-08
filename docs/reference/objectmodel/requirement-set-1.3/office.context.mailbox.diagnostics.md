---
title: Office.--の要件セット1.3
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 46f58490d3f16f2a1bcd4caf0a21734861477e9f
ms.sourcegitcommit: dc78ee2a89fe3d4cd6f748be1eec9081c1077502
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2019
ms.locfileid: "36231278"
---
# <a name="diagnostics"></a><span data-ttu-id="2a445-102">診断</span><span class="sxs-lookup"><span data-stu-id="2a445-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="2a445-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="2a445-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="2a445-104">Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="2a445-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a445-105">要件</span><span class="sxs-lookup"><span data-stu-id="2a445-105">Requirements</span></span>

|<span data-ttu-id="2a445-106">要件</span><span class="sxs-lookup"><span data-stu-id="2a445-106">Requirement</span></span>| <span data-ttu-id="2a445-107">値</span><span class="sxs-lookup"><span data-stu-id="2a445-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a445-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2a445-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2a445-109">1.0</span><span class="sxs-lookup"><span data-stu-id="2a445-109">1.0</span></span>|
|[<span data-ttu-id="2a445-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2a445-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2a445-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a445-111">ReadItem</span></span>|
|[<span data-ttu-id="2a445-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2a445-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2a445-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2a445-113">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="2a445-114">Members</span><span class="sxs-lookup"><span data-stu-id="2a445-114">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="2a445-115">hostName: String</span><span class="sxs-lookup"><span data-stu-id="2a445-115">hostName: String</span></span>

<span data-ttu-id="2a445-116">ホスト アプリケーションの名前を表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="2a445-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="2a445-117">文字列は、値 `Outlook`、`OutlookIOS`、`OutlookWebApp` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="2a445-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="2a445-118">型</span><span class="sxs-lookup"><span data-stu-id="2a445-118">Type</span></span>

*   <span data-ttu-id="2a445-119">String</span><span class="sxs-lookup"><span data-stu-id="2a445-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a445-120">要件</span><span class="sxs-lookup"><span data-stu-id="2a445-120">Requirements</span></span>

|<span data-ttu-id="2a445-121">要件</span><span class="sxs-lookup"><span data-stu-id="2a445-121">Requirement</span></span>| <span data-ttu-id="2a445-122">値</span><span class="sxs-lookup"><span data-stu-id="2a445-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a445-123">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2a445-123">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2a445-124">1.0</span><span class="sxs-lookup"><span data-stu-id="2a445-124">1.0</span></span>|
|[<span data-ttu-id="2a445-125">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2a445-125">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2a445-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a445-126">ReadItem</span></span>|
|[<span data-ttu-id="2a445-127">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2a445-127">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2a445-128">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2a445-128">Compose or Read</span></span>|

#### <a name="hostversion-string"></a><span data-ttu-id="2a445-129">hostVersion: String</span><span class="sxs-lookup"><span data-stu-id="2a445-129">hostVersion: String</span></span>

<span data-ttu-id="2a445-130">ホスト アプリケーションまたは Exchange Server のバージョンを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="2a445-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="2a445-131">メールアドインが Outlook デスクトップクライアントまたは iOS で実行されている場合、 `hostVersion`このプロパティはホストアプリケーションのバージョン (outlook) を返します。</span><span class="sxs-lookup"><span data-stu-id="2a445-131">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="2a445-132">Web 上の Outlook では、このプロパティは Exchange サーバーのバージョンを返します。</span><span class="sxs-lookup"><span data-stu-id="2a445-132">In Outlook on the web, the property returns the version of the Exchange Server.</span></span> <span data-ttu-id="2a445-133">例として、"15.0.468.0" という文字列があります。</span><span class="sxs-lookup"><span data-stu-id="2a445-133">An example is the string "15.0.468.0".</span></span>

##### <a name="type"></a><span data-ttu-id="2a445-134">型</span><span class="sxs-lookup"><span data-stu-id="2a445-134">Type</span></span>

*   <span data-ttu-id="2a445-135">String</span><span class="sxs-lookup"><span data-stu-id="2a445-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a445-136">要件</span><span class="sxs-lookup"><span data-stu-id="2a445-136">Requirements</span></span>

|<span data-ttu-id="2a445-137">要件</span><span class="sxs-lookup"><span data-stu-id="2a445-137">Requirement</span></span>| <span data-ttu-id="2a445-138">値</span><span class="sxs-lookup"><span data-stu-id="2a445-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a445-139">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2a445-139">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2a445-140">1.0</span><span class="sxs-lookup"><span data-stu-id="2a445-140">1.0</span></span>|
|[<span data-ttu-id="2a445-141">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2a445-141">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2a445-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a445-142">ReadItem</span></span>|
|[<span data-ttu-id="2a445-143">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2a445-143">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2a445-144">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2a445-144">Compose or Read</span></span>|

#### <a name="owaview-string"></a><span data-ttu-id="2a445-145">OWAView: String</span><span class="sxs-lookup"><span data-stu-id="2a445-145">OWAView: String</span></span>

<span data-ttu-id="2a445-146">Web 上の Outlook の現在のビューを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="2a445-146">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="2a445-147">返される文字列は、値 `OneColumn`、`TwoColumns`、または `ThreeColumns` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="2a445-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="2a445-148">ホストアプリケーションが web 上の Outlook ではない場合、このプロパティにアクセスする`undefined`と、になります。</span><span class="sxs-lookup"><span data-stu-id="2a445-148">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="2a445-149">Outlook on the web には、画面とウィンドウの幅、および表示できる列の数に対応する3つのビューがあります。</span><span class="sxs-lookup"><span data-stu-id="2a445-149">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="2a445-150">画面幅が狭い場合に表示される `OneColumn`。</span><span class="sxs-lookup"><span data-stu-id="2a445-150">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="2a445-151">Outlook on the web では、スマートフォンの画面全体でこのような単一の列のレイアウトを使用します。</span><span class="sxs-lookup"><span data-stu-id="2a445-151">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="2a445-152">画面幅がやや広い場合に表示される `TwoColumns`。</span><span class="sxs-lookup"><span data-stu-id="2a445-152">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="2a445-153">Web 上の Outlook は、ほとんどのタブレットでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="2a445-153">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="2a445-154">画面幅が広い場合に表示される `ThreeColumns`。</span><span class="sxs-lookup"><span data-stu-id="2a445-154">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="2a445-155">たとえば、Outlook on the web では、このビューをデスクトップコンピューターの全画面表示ウィンドウで使用します。</span><span class="sxs-lookup"><span data-stu-id="2a445-155">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="2a445-156">型</span><span class="sxs-lookup"><span data-stu-id="2a445-156">Type</span></span>

*   <span data-ttu-id="2a445-157">String</span><span class="sxs-lookup"><span data-stu-id="2a445-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a445-158">要件</span><span class="sxs-lookup"><span data-stu-id="2a445-158">Requirements</span></span>

|<span data-ttu-id="2a445-159">要件</span><span class="sxs-lookup"><span data-stu-id="2a445-159">Requirement</span></span>| <span data-ttu-id="2a445-160">値</span><span class="sxs-lookup"><span data-stu-id="2a445-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a445-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2a445-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2a445-162">1.0</span><span class="sxs-lookup"><span data-stu-id="2a445-162">1.0</span></span>|
|[<span data-ttu-id="2a445-163">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2a445-163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2a445-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a445-164">ReadItem</span></span>|
|[<span data-ttu-id="2a445-165">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2a445-165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2a445-166">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2a445-166">Compose or Read</span></span>|
