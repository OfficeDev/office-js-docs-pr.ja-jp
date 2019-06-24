---
title: Office の設定-プレビュー要件セット
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 86ee4093dbfb35a5306938da27b61eb6e8936792
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127577"
---
# <a name="diagnostics"></a><span data-ttu-id="f54fd-102">診断</span><span class="sxs-lookup"><span data-stu-id="f54fd-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="f54fd-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="f54fd-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="f54fd-104">Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="f54fd-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f54fd-105">要件</span><span class="sxs-lookup"><span data-stu-id="f54fd-105">Requirements</span></span>

|<span data-ttu-id="f54fd-106">要件</span><span class="sxs-lookup"><span data-stu-id="f54fd-106">Requirement</span></span>| <span data-ttu-id="f54fd-107">値</span><span class="sxs-lookup"><span data-stu-id="f54fd-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f54fd-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f54fd-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f54fd-109">1.0</span><span class="sxs-lookup"><span data-stu-id="f54fd-109">1.0</span></span>|
|[<span data-ttu-id="f54fd-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f54fd-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f54fd-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f54fd-111">ReadItem</span></span>|
|[<span data-ttu-id="f54fd-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f54fd-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f54fd-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f54fd-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f54fd-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="f54fd-114">Members and methods</span></span>

| <span data-ttu-id="f54fd-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="f54fd-115">Member</span></span> | <span data-ttu-id="f54fd-116">種類</span><span class="sxs-lookup"><span data-stu-id="f54fd-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f54fd-117">名</span><span class="sxs-lookup"><span data-stu-id="f54fd-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="f54fd-118">Member</span><span class="sxs-lookup"><span data-stu-id="f54fd-118">Member</span></span> |
| [<span data-ttu-id="f54fd-119">上 diagnostics.hostversion</span><span class="sxs-lookup"><span data-stu-id="f54fd-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="f54fd-120">Member</span><span class="sxs-lookup"><span data-stu-id="f54fd-120">Member</span></span> |
| [<span data-ttu-id="f54fd-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="f54fd-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="f54fd-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="f54fd-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="f54fd-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="f54fd-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="f54fd-124">hostName: String</span><span class="sxs-lookup"><span data-stu-id="f54fd-124">hostName: String</span></span>

<span data-ttu-id="f54fd-125">ホスト アプリケーションの名前を表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="f54fd-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="f54fd-126">文字列は、値 `Outlook`、`Mac Outlook`、`OutlookIOS`、または `OutlookWebApp` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="f54fd-126">A string that can be one of the following values: `Outlook`, `Mac Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="f54fd-127">型</span><span class="sxs-lookup"><span data-stu-id="f54fd-127">Type</span></span>

*   <span data-ttu-id="f54fd-128">String</span><span class="sxs-lookup"><span data-stu-id="f54fd-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f54fd-129">要件</span><span class="sxs-lookup"><span data-stu-id="f54fd-129">Requirements</span></span>

|<span data-ttu-id="f54fd-130">要件</span><span class="sxs-lookup"><span data-stu-id="f54fd-130">Requirement</span></span>| <span data-ttu-id="f54fd-131">値</span><span class="sxs-lookup"><span data-stu-id="f54fd-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="f54fd-132">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f54fd-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f54fd-133">1.0</span><span class="sxs-lookup"><span data-stu-id="f54fd-133">1.0</span></span>|
|[<span data-ttu-id="f54fd-134">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f54fd-134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f54fd-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f54fd-135">ReadItem</span></span>|
|[<span data-ttu-id="f54fd-136">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f54fd-136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f54fd-137">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f54fd-137">Compose or Read</span></span>|

---
---

#### <a name="hostversion-string"></a><span data-ttu-id="f54fd-138">hostVersion: String</span><span class="sxs-lookup"><span data-stu-id="f54fd-138">hostVersion: String</span></span>

<span data-ttu-id="f54fd-139">ホスト アプリケーションまたは Exchange Server のバージョンを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="f54fd-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="f54fd-140">メールアドインが Outlook デスクトップクライアント上または iOS 上で実行されている場合`hostVersion` 、このプロパティはホストアプリケーションのバージョン (outlook) を返します。</span><span class="sxs-lookup"><span data-stu-id="f54fd-140">If the mail add-in is running on the Outlook desktop client or on iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="f54fd-141">Web 上の Outlook では、このプロパティは Exchange サーバーのバージョンを返します。</span><span class="sxs-lookup"><span data-stu-id="f54fd-141">In Outlook on the web, the property returns the version of the Exchange Server.</span></span> <span data-ttu-id="f54fd-142">たとえば、文字列 `15.0.468.0` です。</span><span class="sxs-lookup"><span data-stu-id="f54fd-142">An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="f54fd-143">型</span><span class="sxs-lookup"><span data-stu-id="f54fd-143">Type</span></span>

*   <span data-ttu-id="f54fd-144">String</span><span class="sxs-lookup"><span data-stu-id="f54fd-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f54fd-145">要件</span><span class="sxs-lookup"><span data-stu-id="f54fd-145">Requirements</span></span>

|<span data-ttu-id="f54fd-146">要件</span><span class="sxs-lookup"><span data-stu-id="f54fd-146">Requirement</span></span>| <span data-ttu-id="f54fd-147">値</span><span class="sxs-lookup"><span data-stu-id="f54fd-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="f54fd-148">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f54fd-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f54fd-149">1.0</span><span class="sxs-lookup"><span data-stu-id="f54fd-149">1.0</span></span>|
|[<span data-ttu-id="f54fd-150">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f54fd-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f54fd-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f54fd-151">ReadItem</span></span>|
|[<span data-ttu-id="f54fd-152">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f54fd-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f54fd-153">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f54fd-153">Compose or Read</span></span>|

---
---

#### <a name="owaview-string"></a><span data-ttu-id="f54fd-154">OWAView: String</span><span class="sxs-lookup"><span data-stu-id="f54fd-154">OWAView: String</span></span>

<span data-ttu-id="f54fd-155">Web 上の Outlook の現在のビューを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="f54fd-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="f54fd-156">返される文字列は、値 `OneColumn`、`TwoColumns`、または `ThreeColumns` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="f54fd-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="f54fd-157">ホストアプリケーションが web 上の Outlook ではない場合、このプロパティにアクセスする`undefined`と、になります。</span><span class="sxs-lookup"><span data-stu-id="f54fd-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="f54fd-158">Outlook on the web には、画面とウィンドウの幅、および表示できる列の数に対応する3つのビューがあります。</span><span class="sxs-lookup"><span data-stu-id="f54fd-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="f54fd-159">画面幅が狭い場合に表示される `OneColumn`。</span><span class="sxs-lookup"><span data-stu-id="f54fd-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="f54fd-160">Outlook on the web では、スマートフォンの画面全体でこのような単一の列のレイアウトを使用します。</span><span class="sxs-lookup"><span data-stu-id="f54fd-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="f54fd-161">画面幅がやや広い場合に表示される `TwoColumns`。</span><span class="sxs-lookup"><span data-stu-id="f54fd-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="f54fd-162">Web 上の Outlook は、ほとんどのタブレットでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="f54fd-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="f54fd-163">画面幅が広い場合に表示される `ThreeColumns`。</span><span class="sxs-lookup"><span data-stu-id="f54fd-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="f54fd-164">たとえば、Outlook on the web では、このビューをデスクトップコンピューターの全画面表示ウィンドウで使用します。</span><span class="sxs-lookup"><span data-stu-id="f54fd-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="f54fd-165">型</span><span class="sxs-lookup"><span data-stu-id="f54fd-165">Type</span></span>

*   <span data-ttu-id="f54fd-166">String</span><span class="sxs-lookup"><span data-stu-id="f54fd-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f54fd-167">要件</span><span class="sxs-lookup"><span data-stu-id="f54fd-167">Requirements</span></span>

|<span data-ttu-id="f54fd-168">要件</span><span class="sxs-lookup"><span data-stu-id="f54fd-168">Requirement</span></span>| <span data-ttu-id="f54fd-169">値</span><span class="sxs-lookup"><span data-stu-id="f54fd-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="f54fd-170">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f54fd-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f54fd-171">1.0</span><span class="sxs-lookup"><span data-stu-id="f54fd-171">1.0</span></span>|
|[<span data-ttu-id="f54fd-172">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f54fd-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f54fd-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f54fd-173">ReadItem</span></span>|
|[<span data-ttu-id="f54fd-174">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f54fd-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f54fd-175">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f54fd-175">Compose or Read</span></span>|
