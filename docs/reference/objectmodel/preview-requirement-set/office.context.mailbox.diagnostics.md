---
title: Office の設定-プレビュー要件セット
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 492e292737417854adfaf98feb2b67788933d874
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629203"
---
# <a name="diagnostics"></a><span data-ttu-id="39fb9-102">診断</span><span class="sxs-lookup"><span data-stu-id="39fb9-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="39fb9-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="39fb9-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="39fb9-104">Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="39fb9-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="39fb9-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="39fb9-105">Requirements</span></span>

|<span data-ttu-id="39fb9-106">要件</span><span class="sxs-lookup"><span data-stu-id="39fb9-106">Requirement</span></span>| <span data-ttu-id="39fb9-107">値</span><span class="sxs-lookup"><span data-stu-id="39fb9-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="39fb9-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="39fb9-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39fb9-109">1.0</span><span class="sxs-lookup"><span data-stu-id="39fb9-109">1.0</span></span>|
|[<span data-ttu-id="39fb9-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="39fb9-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39fb9-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39fb9-111">ReadItem</span></span>|
|[<span data-ttu-id="39fb9-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="39fb9-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39fb9-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="39fb9-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="39fb9-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="39fb9-114">Properties</span></span>

| <span data-ttu-id="39fb9-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="39fb9-115">Property</span></span> | <span data-ttu-id="39fb9-116">最小値</span><span class="sxs-lookup"><span data-stu-id="39fb9-116">Minimum</span></span><br><span data-ttu-id="39fb9-117">アクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="39fb9-117">permission level</span></span> | <span data-ttu-id="39fb9-118">モード</span><span class="sxs-lookup"><span data-stu-id="39fb9-118">Modes</span></span> | <span data-ttu-id="39fb9-119">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="39fb9-119">Return type</span></span> | <span data-ttu-id="39fb9-120">最小値</span><span class="sxs-lookup"><span data-stu-id="39fb9-120">Minimum</span></span><br><span data-ttu-id="39fb9-121">要件セット</span><span class="sxs-lookup"><span data-stu-id="39fb9-121">requirement set</span></span> |
|---|---|---|---|---|
| [<span data-ttu-id="39fb9-122">名</span><span class="sxs-lookup"><span data-stu-id="39fb9-122">hostName</span></span>](#hostname-string) | <span data-ttu-id="39fb9-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39fb9-123">ReadItem</span></span> | <span data-ttu-id="39fb9-124">作成</span><span class="sxs-lookup"><span data-stu-id="39fb9-124">Compose</span></span><br><span data-ttu-id="39fb9-125">読み取り</span><span class="sxs-lookup"><span data-stu-id="39fb9-125">Read</span></span> | <span data-ttu-id="39fb9-126">String</span><span class="sxs-lookup"><span data-stu-id="39fb9-126">String</span></span> | <span data-ttu-id="39fb9-127">1.0</span><span class="sxs-lookup"><span data-stu-id="39fb9-127">1.0</span></span> |
| [<span data-ttu-id="39fb9-128">上 diagnostics.hostversion</span><span class="sxs-lookup"><span data-stu-id="39fb9-128">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="39fb9-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39fb9-129">ReadItem</span></span> | <span data-ttu-id="39fb9-130">作成</span><span class="sxs-lookup"><span data-stu-id="39fb9-130">Compose</span></span><br><span data-ttu-id="39fb9-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="39fb9-131">Read</span></span> | <span data-ttu-id="39fb9-132">String</span><span class="sxs-lookup"><span data-stu-id="39fb9-132">String</span></span> | <span data-ttu-id="39fb9-133">1.0</span><span class="sxs-lookup"><span data-stu-id="39fb9-133">1.0</span></span> |
| [<span data-ttu-id="39fb9-134">OWAView</span><span class="sxs-lookup"><span data-stu-id="39fb9-134">OWAView</span></span>](#owaview-string) | <span data-ttu-id="39fb9-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39fb9-135">ReadItem</span></span> | <span data-ttu-id="39fb9-136">作成</span><span class="sxs-lookup"><span data-stu-id="39fb9-136">Compose</span></span><br><span data-ttu-id="39fb9-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="39fb9-137">Read</span></span> | <span data-ttu-id="39fb9-138">String</span><span class="sxs-lookup"><span data-stu-id="39fb9-138">String</span></span> | <span data-ttu-id="39fb9-139">1.0</span><span class="sxs-lookup"><span data-stu-id="39fb9-139">1.0</span></span> |

## <a name="property-details"></a><span data-ttu-id="39fb9-140">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="39fb9-140">Property details</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="39fb9-141">hostName: String</span><span class="sxs-lookup"><span data-stu-id="39fb9-141">hostName: String</span></span>

<span data-ttu-id="39fb9-142">ホスト アプリケーションの名前を表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="39fb9-142">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="39fb9-143">文字列は、値 `Outlook`、`OutlookWebApp`、`OutlookIOS`、または `OutlookAndroid` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="39fb9-143">A string that can be one of the following values: `Outlook`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.</span></span>

> [!NOTE]
> <span data-ttu-id="39fb9-144">この`Outlook`値は、デスクトップクライアント (つまり Windows と Mac) の Outlook に対して返されます。</span><span class="sxs-lookup"><span data-stu-id="39fb9-144">The `Outlook` value is returned for Outlook on desktop clients (i.e., Windows and Mac).</span></span>

##### <a name="type"></a><span data-ttu-id="39fb9-145">型</span><span class="sxs-lookup"><span data-stu-id="39fb9-145">Type</span></span>

*   <span data-ttu-id="39fb9-146">String</span><span class="sxs-lookup"><span data-stu-id="39fb9-146">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="39fb9-147">要件</span><span class="sxs-lookup"><span data-stu-id="39fb9-147">Requirements</span></span>

|<span data-ttu-id="39fb9-148">要件</span><span class="sxs-lookup"><span data-stu-id="39fb9-148">Requirement</span></span>| <span data-ttu-id="39fb9-149">値</span><span class="sxs-lookup"><span data-stu-id="39fb9-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="39fb9-150">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="39fb9-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39fb9-151">1.0</span><span class="sxs-lookup"><span data-stu-id="39fb9-151">1.0</span></span>|
|[<span data-ttu-id="39fb9-152">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="39fb9-152">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39fb9-153">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39fb9-153">ReadItem</span></span>|
|[<span data-ttu-id="39fb9-154">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="39fb9-154">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39fb9-155">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="39fb9-155">Compose or Read</span></span>|

<br>

---
---

#### <a name="hostversion-string"></a><span data-ttu-id="39fb9-156">hostVersion: String</span><span class="sxs-lookup"><span data-stu-id="39fb9-156">hostVersion: String</span></span>

<span data-ttu-id="39fb9-157">ホストアプリケーションまたは Exchange サーバー (例: "15.0.468.0") のバージョンを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="39fb9-157">Gets a string that represents the version of either the host application or the Exchange Server (e.g., "15.0.468.0").</span></span>

<span data-ttu-id="39fb9-158">メールアドインが Outlook デスクトップまたはモバイルクライアント上で実行されている場合`hostVersion` 、このプロパティはホストアプリケーションのバージョン (outlook) を返します。</span><span class="sxs-lookup"><span data-stu-id="39fb9-158">If the mail add-in is running on an Outlook desktop or mobile client, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="39fb9-159">Web 上の Outlook では、このプロパティは Exchange サーバーのバージョンを返します。</span><span class="sxs-lookup"><span data-stu-id="39fb9-159">In Outlook on the web, the property returns the version of the Exchange Server.</span></span>

##### <a name="type"></a><span data-ttu-id="39fb9-160">型</span><span class="sxs-lookup"><span data-stu-id="39fb9-160">Type</span></span>

*   <span data-ttu-id="39fb9-161">String</span><span class="sxs-lookup"><span data-stu-id="39fb9-161">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="39fb9-162">要件</span><span class="sxs-lookup"><span data-stu-id="39fb9-162">Requirements</span></span>

|<span data-ttu-id="39fb9-163">要件</span><span class="sxs-lookup"><span data-stu-id="39fb9-163">Requirement</span></span>| <span data-ttu-id="39fb9-164">値</span><span class="sxs-lookup"><span data-stu-id="39fb9-164">Value</span></span>|
|---|---|
|[<span data-ttu-id="39fb9-165">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="39fb9-165">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39fb9-166">1.0</span><span class="sxs-lookup"><span data-stu-id="39fb9-166">1.0</span></span>|
|[<span data-ttu-id="39fb9-167">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="39fb9-167">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39fb9-168">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39fb9-168">ReadItem</span></span>|
|[<span data-ttu-id="39fb9-169">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="39fb9-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39fb9-170">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="39fb9-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="owaview-string"></a><span data-ttu-id="39fb9-171">OWAView: String</span><span class="sxs-lookup"><span data-stu-id="39fb9-171">OWAView: String</span></span>

<span data-ttu-id="39fb9-172">Web 上の Outlook の現在のビューを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="39fb9-172">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="39fb9-173">返される文字列は、値 `OneColumn`、`TwoColumns`、または `ThreeColumns` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="39fb9-173">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="39fb9-174">ホストアプリケーションが web 上の Outlook ではない場合、このプロパティにアクセスする`undefined`と、になります。</span><span class="sxs-lookup"><span data-stu-id="39fb9-174">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="39fb9-175">Outlook on the web には、画面とウィンドウの幅、および表示できる列の数に対応する3つのビューがあります。</span><span class="sxs-lookup"><span data-stu-id="39fb9-175">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="39fb9-176">画面幅が狭い場合に表示される `OneColumn`。</span><span class="sxs-lookup"><span data-stu-id="39fb9-176">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="39fb9-177">Outlook on the web では、スマートフォンの画面全体でこのような単一の列のレイアウトを使用します。</span><span class="sxs-lookup"><span data-stu-id="39fb9-177">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="39fb9-178">画面幅がやや広い場合に表示される `TwoColumns`。</span><span class="sxs-lookup"><span data-stu-id="39fb9-178">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="39fb9-179">Web 上の Outlook は、ほとんどのタブレットでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="39fb9-179">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="39fb9-180">画面幅が広い場合に表示される `ThreeColumns`。</span><span class="sxs-lookup"><span data-stu-id="39fb9-180">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="39fb9-181">たとえば、Outlook on the web では、このビューをデスクトップコンピューターの全画面表示ウィンドウで使用します。</span><span class="sxs-lookup"><span data-stu-id="39fb9-181">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="39fb9-182">型</span><span class="sxs-lookup"><span data-stu-id="39fb9-182">Type</span></span>

*   <span data-ttu-id="39fb9-183">String</span><span class="sxs-lookup"><span data-stu-id="39fb9-183">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="39fb9-184">要件</span><span class="sxs-lookup"><span data-stu-id="39fb9-184">Requirements</span></span>

|<span data-ttu-id="39fb9-185">要件</span><span class="sxs-lookup"><span data-stu-id="39fb9-185">Requirement</span></span>| <span data-ttu-id="39fb9-186">値</span><span class="sxs-lookup"><span data-stu-id="39fb9-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="39fb9-187">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="39fb9-187">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="39fb9-188">1.0</span><span class="sxs-lookup"><span data-stu-id="39fb9-188">1.0</span></span>|
|[<span data-ttu-id="39fb9-189">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="39fb9-189">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39fb9-190">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39fb9-190">ReadItem</span></span>|
|[<span data-ttu-id="39fb9-191">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="39fb9-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39fb9-192">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="39fb9-192">Compose or Read</span></span>|
