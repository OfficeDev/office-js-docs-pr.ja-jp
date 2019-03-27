---
title: Office.--の要件セット1.5
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 7e554217831f2739ead3a0a90bd41b7d72e7b2d1
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871116"
---
# <a name="diagnostics"></a><span data-ttu-id="1216f-102">診断</span><span class="sxs-lookup"><span data-stu-id="1216f-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="1216f-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="1216f-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="1216f-104">Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="1216f-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1216f-105">要件</span><span class="sxs-lookup"><span data-stu-id="1216f-105">Requirements</span></span>

|<span data-ttu-id="1216f-106">要件</span><span class="sxs-lookup"><span data-stu-id="1216f-106">Requirement</span></span>| <span data-ttu-id="1216f-107">値</span><span class="sxs-lookup"><span data-stu-id="1216f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="1216f-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1216f-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1216f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="1216f-109">1.0</span></span>|
|[<span data-ttu-id="1216f-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1216f-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1216f-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1216f-111">ReadItem</span></span>|
|[<span data-ttu-id="1216f-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1216f-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1216f-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1216f-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1216f-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="1216f-114">Members and methods</span></span>

| <span data-ttu-id="1216f-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="1216f-115">Member</span></span> | <span data-ttu-id="1216f-116">種類</span><span class="sxs-lookup"><span data-stu-id="1216f-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1216f-117">名</span><span class="sxs-lookup"><span data-stu-id="1216f-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="1216f-118">Member</span><span class="sxs-lookup"><span data-stu-id="1216f-118">Member</span></span> |
| [<span data-ttu-id="1216f-119">上 diagnostics.hostversion</span><span class="sxs-lookup"><span data-stu-id="1216f-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="1216f-120">Member</span><span class="sxs-lookup"><span data-stu-id="1216f-120">Member</span></span> |
| [<span data-ttu-id="1216f-121">owaview</span><span class="sxs-lookup"><span data-stu-id="1216f-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="1216f-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="1216f-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="1216f-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="1216f-123">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="1216f-124">hostName :String</span><span class="sxs-lookup"><span data-stu-id="1216f-124">hostName :String</span></span>

<span data-ttu-id="1216f-125">ホスト アプリケーションの名前を表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="1216f-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="1216f-126">文字列は、値 `Outlook`、`OutlookIOS`、`OutlookWebApp` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="1216f-126">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="1216f-127">型</span><span class="sxs-lookup"><span data-stu-id="1216f-127">Type</span></span>

*   <span data-ttu-id="1216f-128">String</span><span class="sxs-lookup"><span data-stu-id="1216f-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1216f-129">要件</span><span class="sxs-lookup"><span data-stu-id="1216f-129">Requirements</span></span>

|<span data-ttu-id="1216f-130">要件</span><span class="sxs-lookup"><span data-stu-id="1216f-130">Requirement</span></span>| <span data-ttu-id="1216f-131">値</span><span class="sxs-lookup"><span data-stu-id="1216f-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="1216f-132">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1216f-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1216f-133">1.0</span><span class="sxs-lookup"><span data-stu-id="1216f-133">1.0</span></span>|
|[<span data-ttu-id="1216f-134">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1216f-134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1216f-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1216f-135">ReadItem</span></span>|
|[<span data-ttu-id="1216f-136">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1216f-136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1216f-137">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1216f-137">Compose or Read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="1216f-138">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="1216f-138">hostVersion :String</span></span>

<span data-ttu-id="1216f-139">ホスト アプリケーションまたは Exchange Server のバージョンを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="1216f-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="1216f-p101">メール アドインを Outlook デスクトップ クライアントまたは Outlook for iOS で実行している場合、`hostVersion` プロパティは、ホスト アプリケーションである Outlook のバージョンを返します。Outlook Web App では、プロパティは、Exchange Server のバージョンを返します。たとえば、文字列 `15.0.468.0` です。</span><span class="sxs-lookup"><span data-stu-id="1216f-p101">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="1216f-143">型</span><span class="sxs-lookup"><span data-stu-id="1216f-143">Type</span></span>

*   <span data-ttu-id="1216f-144">String</span><span class="sxs-lookup"><span data-stu-id="1216f-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1216f-145">要件</span><span class="sxs-lookup"><span data-stu-id="1216f-145">Requirements</span></span>

|<span data-ttu-id="1216f-146">要件</span><span class="sxs-lookup"><span data-stu-id="1216f-146">Requirement</span></span>| <span data-ttu-id="1216f-147">値</span><span class="sxs-lookup"><span data-stu-id="1216f-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="1216f-148">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1216f-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1216f-149">1.0</span><span class="sxs-lookup"><span data-stu-id="1216f-149">1.0</span></span>|
|[<span data-ttu-id="1216f-150">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1216f-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1216f-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1216f-151">ReadItem</span></span>|
|[<span data-ttu-id="1216f-152">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1216f-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1216f-153">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1216f-153">Compose or Read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="1216f-154">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="1216f-154">OWAView :String</span></span>

<span data-ttu-id="1216f-155">Outlook Web App の現在のビューを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="1216f-155">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="1216f-156">返される文字列は、値 `OneColumn`、`TwoColumns`、または `ThreeColumns` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="1216f-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="1216f-157">ホスト アプリケーションが Outlook Web App ではない場合、このプロパティにアクセスすると `undefined` が返されます。</span><span class="sxs-lookup"><span data-stu-id="1216f-157">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="1216f-158">Outlook Web App には、画面とウィンドウの幅、および表示可能な列数に応じて 3 つのビューがあります。</span><span class="sxs-lookup"><span data-stu-id="1216f-158">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="1216f-p102">画面幅が狭い場合に表示される `OneColumn`。Outlook Web App は、この単一列レイアウトを使用してスマートフォンの画面全体への表示を行います。</span><span class="sxs-lookup"><span data-stu-id="1216f-p102">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="1216f-p103">画面幅がやや広い場合に表示される `TwoColumns`。Outlook Web App は、ほとんどのタブレットでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="1216f-p103">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="1216f-p104">画面幅が広い場合に表示される `ThreeColumns`。Outlook Web App は、デスクトップ コンピューターのフル スクリーン ウィンドウなどでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="1216f-p104">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="1216f-165">型</span><span class="sxs-lookup"><span data-stu-id="1216f-165">Type</span></span>

*   <span data-ttu-id="1216f-166">String</span><span class="sxs-lookup"><span data-stu-id="1216f-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1216f-167">要件</span><span class="sxs-lookup"><span data-stu-id="1216f-167">Requirements</span></span>

|<span data-ttu-id="1216f-168">要件</span><span class="sxs-lookup"><span data-stu-id="1216f-168">Requirement</span></span>| <span data-ttu-id="1216f-169">値</span><span class="sxs-lookup"><span data-stu-id="1216f-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="1216f-170">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1216f-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1216f-171">1.0</span><span class="sxs-lookup"><span data-stu-id="1216f-171">1.0</span></span>|
|[<span data-ttu-id="1216f-172">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1216f-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1216f-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1216f-173">ReadItem</span></span>|
|[<span data-ttu-id="1216f-174">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1216f-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1216f-175">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1216f-175">Compose or Read</span></span>|
