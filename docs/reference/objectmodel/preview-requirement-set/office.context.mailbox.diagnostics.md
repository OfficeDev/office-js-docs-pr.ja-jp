---
title: Office の設定-プレビュー要件セット
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 0d0283edc0b9c331afb15a9b485c8f1447d5e8e5
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268531"
---
# <a name="diagnostics"></a><span data-ttu-id="ac928-102">診断</span><span class="sxs-lookup"><span data-stu-id="ac928-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="ac928-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="ac928-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="ac928-104">Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="ac928-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ac928-105">要件</span><span class="sxs-lookup"><span data-stu-id="ac928-105">Requirements</span></span>

|<span data-ttu-id="ac928-106">要件</span><span class="sxs-lookup"><span data-stu-id="ac928-106">Requirement</span></span>| <span data-ttu-id="ac928-107">値</span><span class="sxs-lookup"><span data-stu-id="ac928-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac928-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ac928-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ac928-109">1.0</span><span class="sxs-lookup"><span data-stu-id="ac928-109">1.0</span></span>|
|[<span data-ttu-id="ac928-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ac928-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ac928-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ac928-111">ReadItem</span></span>|
|[<span data-ttu-id="ac928-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ac928-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ac928-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ac928-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ac928-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="ac928-114">Members and methods</span></span>

| <span data-ttu-id="ac928-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="ac928-115">Member</span></span> | <span data-ttu-id="ac928-116">種類</span><span class="sxs-lookup"><span data-stu-id="ac928-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ac928-117">名</span><span class="sxs-lookup"><span data-stu-id="ac928-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="ac928-118">Member</span><span class="sxs-lookup"><span data-stu-id="ac928-118">Member</span></span> |
| [<span data-ttu-id="ac928-119">上 diagnostics.hostversion</span><span class="sxs-lookup"><span data-stu-id="ac928-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="ac928-120">Member</span><span class="sxs-lookup"><span data-stu-id="ac928-120">Member</span></span> |
| [<span data-ttu-id="ac928-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="ac928-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="ac928-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="ac928-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="ac928-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="ac928-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="ac928-124">hostName: String</span><span class="sxs-lookup"><span data-stu-id="ac928-124">hostName: String</span></span>

<span data-ttu-id="ac928-125">ホスト アプリケーションの名前を表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="ac928-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="ac928-126">文字列は、値 `Outlook`、`OutlookWebApp`、`OutlookIOS`、または `OutlookAndroid` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="ac928-126">A string that can be one of the following values: `Outlook`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.</span></span>

> [!NOTE]
> <span data-ttu-id="ac928-127">この`Outlook`値は、デスクトップクライアント (つまり Windows と Mac) の Outlook に対して返されます。</span><span class="sxs-lookup"><span data-stu-id="ac928-127">The `Outlook` value is returned for Outlook on desktop clients (i.e., Windows and Mac).</span></span>

##### <a name="type"></a><span data-ttu-id="ac928-128">型</span><span class="sxs-lookup"><span data-stu-id="ac928-128">Type</span></span>

*   <span data-ttu-id="ac928-129">String</span><span class="sxs-lookup"><span data-stu-id="ac928-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ac928-130">要件</span><span class="sxs-lookup"><span data-stu-id="ac928-130">Requirements</span></span>

|<span data-ttu-id="ac928-131">要件</span><span class="sxs-lookup"><span data-stu-id="ac928-131">Requirement</span></span>| <span data-ttu-id="ac928-132">値</span><span class="sxs-lookup"><span data-stu-id="ac928-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac928-133">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ac928-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ac928-134">1.0</span><span class="sxs-lookup"><span data-stu-id="ac928-134">1.0</span></span>|
|[<span data-ttu-id="ac928-135">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ac928-135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ac928-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ac928-136">ReadItem</span></span>|
|[<span data-ttu-id="ac928-137">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ac928-137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ac928-138">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ac928-138">Compose or Read</span></span>|

---
---

#### <a name="hostversion-string"></a><span data-ttu-id="ac928-139">hostVersion: String</span><span class="sxs-lookup"><span data-stu-id="ac928-139">hostVersion: String</span></span>

<span data-ttu-id="ac928-140">ホストアプリケーションまたは Exchange サーバー (例: "15.0.468.0") のバージョンを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="ac928-140">Gets a string that represents the version of either the host application or the Exchange Server (e.g., "15.0.468.0").</span></span>

<span data-ttu-id="ac928-141">メールアドインが Outlook デスクトップクライアント上または iOS 上で実行されている場合`hostVersion` 、このプロパティはホストアプリケーションのバージョン (outlook) を返します。</span><span class="sxs-lookup"><span data-stu-id="ac928-141">If the mail add-in is running on the Outlook desktop client or on iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="ac928-142">Web 上の Outlook では、このプロパティは Exchange サーバーのバージョンを返します。</span><span class="sxs-lookup"><span data-stu-id="ac928-142">In Outlook on the web, the property returns the version of the Exchange Server.</span></span>

##### <a name="type"></a><span data-ttu-id="ac928-143">型</span><span class="sxs-lookup"><span data-stu-id="ac928-143">Type</span></span>

*   <span data-ttu-id="ac928-144">String</span><span class="sxs-lookup"><span data-stu-id="ac928-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ac928-145">要件</span><span class="sxs-lookup"><span data-stu-id="ac928-145">Requirements</span></span>

|<span data-ttu-id="ac928-146">要件</span><span class="sxs-lookup"><span data-stu-id="ac928-146">Requirement</span></span>| <span data-ttu-id="ac928-147">値</span><span class="sxs-lookup"><span data-stu-id="ac928-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac928-148">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ac928-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ac928-149">1.0</span><span class="sxs-lookup"><span data-stu-id="ac928-149">1.0</span></span>|
|[<span data-ttu-id="ac928-150">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ac928-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ac928-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ac928-151">ReadItem</span></span>|
|[<span data-ttu-id="ac928-152">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ac928-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ac928-153">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ac928-153">Compose or Read</span></span>|

---
---

#### <a name="owaview-string"></a><span data-ttu-id="ac928-154">OWAView: String</span><span class="sxs-lookup"><span data-stu-id="ac928-154">OWAView: String</span></span>

<span data-ttu-id="ac928-155">Web 上の Outlook の現在のビューを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="ac928-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="ac928-156">返される文字列は、値 `OneColumn`、`TwoColumns`、または `ThreeColumns` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="ac928-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="ac928-157">ホストアプリケーションが web 上の Outlook ではない場合、このプロパティにアクセスする`undefined`と、になります。</span><span class="sxs-lookup"><span data-stu-id="ac928-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="ac928-158">Outlook on the web には、画面とウィンドウの幅、および表示できる列の数に対応する3つのビューがあります。</span><span class="sxs-lookup"><span data-stu-id="ac928-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="ac928-159">画面幅が狭い場合に表示される `OneColumn`。</span><span class="sxs-lookup"><span data-stu-id="ac928-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="ac928-160">Outlook on the web では、スマートフォンの画面全体でこのような単一の列のレイアウトを使用します。</span><span class="sxs-lookup"><span data-stu-id="ac928-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="ac928-161">画面幅がやや広い場合に表示される `TwoColumns`。</span><span class="sxs-lookup"><span data-stu-id="ac928-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="ac928-162">Web 上の Outlook は、ほとんどのタブレットでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="ac928-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="ac928-163">画面幅が広い場合に表示される `ThreeColumns`。</span><span class="sxs-lookup"><span data-stu-id="ac928-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="ac928-164">たとえば、Outlook on the web では、このビューをデスクトップコンピューターの全画面表示ウィンドウで使用します。</span><span class="sxs-lookup"><span data-stu-id="ac928-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="ac928-165">型</span><span class="sxs-lookup"><span data-stu-id="ac928-165">Type</span></span>

*   <span data-ttu-id="ac928-166">String</span><span class="sxs-lookup"><span data-stu-id="ac928-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ac928-167">要件</span><span class="sxs-lookup"><span data-stu-id="ac928-167">Requirements</span></span>

|<span data-ttu-id="ac928-168">要件</span><span class="sxs-lookup"><span data-stu-id="ac928-168">Requirement</span></span>| <span data-ttu-id="ac928-169">値</span><span class="sxs-lookup"><span data-stu-id="ac928-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac928-170">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ac928-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ac928-171">1.0</span><span class="sxs-lookup"><span data-stu-id="ac928-171">1.0</span></span>|
|[<span data-ttu-id="ac928-172">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ac928-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ac928-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ac928-173">ReadItem</span></span>|
|[<span data-ttu-id="ac928-174">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ac928-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ac928-175">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ac928-175">Compose or Read</span></span>|
