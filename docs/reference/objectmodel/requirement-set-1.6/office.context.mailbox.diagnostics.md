---
title: Office.--の要件セット1.6
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 7f6a115f725a238bc87441e0b7c97185240a5c0c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451767"
---
# <a name="diagnostics"></a><span data-ttu-id="04f14-102">診断</span><span class="sxs-lookup"><span data-stu-id="04f14-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="04f14-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="04f14-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="04f14-104">Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="04f14-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="04f14-105">要件</span><span class="sxs-lookup"><span data-stu-id="04f14-105">Requirements</span></span>

|<span data-ttu-id="04f14-106">要件</span><span class="sxs-lookup"><span data-stu-id="04f14-106">Requirement</span></span>| <span data-ttu-id="04f14-107">値</span><span class="sxs-lookup"><span data-stu-id="04f14-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="04f14-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="04f14-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04f14-109">1.0</span><span class="sxs-lookup"><span data-stu-id="04f14-109">1.0</span></span>|
|[<span data-ttu-id="04f14-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="04f14-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04f14-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04f14-111">ReadItem</span></span>|
|[<span data-ttu-id="04f14-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="04f14-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04f14-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="04f14-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="04f14-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="04f14-114">Members and methods</span></span>

| <span data-ttu-id="04f14-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="04f14-115">Member</span></span> | <span data-ttu-id="04f14-116">種類</span><span class="sxs-lookup"><span data-stu-id="04f14-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="04f14-117">名</span><span class="sxs-lookup"><span data-stu-id="04f14-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="04f14-118">Member</span><span class="sxs-lookup"><span data-stu-id="04f14-118">Member</span></span> |
| [<span data-ttu-id="04f14-119">上 diagnostics.hostversion</span><span class="sxs-lookup"><span data-stu-id="04f14-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="04f14-120">Member</span><span class="sxs-lookup"><span data-stu-id="04f14-120">Member</span></span> |
| [<span data-ttu-id="04f14-121">owaview</span><span class="sxs-lookup"><span data-stu-id="04f14-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="04f14-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="04f14-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="04f14-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="04f14-123">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="04f14-124">hostName :String</span><span class="sxs-lookup"><span data-stu-id="04f14-124">hostName :String</span></span>

<span data-ttu-id="04f14-125">ホスト アプリケーションの名前を表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="04f14-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="04f14-126">文字列は、値 `Outlook`、`Mac Outlook`、`OutlookIOS`、または `OutlookWebApp` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="04f14-126">A string that can be one of the following values: `Outlook`, `Mac Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="04f14-127">型</span><span class="sxs-lookup"><span data-stu-id="04f14-127">Type</span></span>

*   <span data-ttu-id="04f14-128">String</span><span class="sxs-lookup"><span data-stu-id="04f14-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="04f14-129">要件</span><span class="sxs-lookup"><span data-stu-id="04f14-129">Requirements</span></span>

|<span data-ttu-id="04f14-130">要件</span><span class="sxs-lookup"><span data-stu-id="04f14-130">Requirement</span></span>| <span data-ttu-id="04f14-131">値</span><span class="sxs-lookup"><span data-stu-id="04f14-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="04f14-132">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="04f14-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04f14-133">1.0</span><span class="sxs-lookup"><span data-stu-id="04f14-133">1.0</span></span>|
|[<span data-ttu-id="04f14-134">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="04f14-134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04f14-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04f14-135">ReadItem</span></span>|
|[<span data-ttu-id="04f14-136">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="04f14-136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04f14-137">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="04f14-137">Compose or Read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="04f14-138">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="04f14-138">hostVersion :String</span></span>

<span data-ttu-id="04f14-139">ホスト アプリケーションまたは Exchange Server のバージョンを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="04f14-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="04f14-p101">メール アドインを Outlook デスクトップ クライアントまたは Outlook for iOS で実行している場合、`hostVersion` プロパティは、ホスト アプリケーションである Outlook のバージョンを返します。Outlook Web App では、プロパティは、Exchange Server のバージョンを返します。たとえば、文字列 `15.0.468.0` です。</span><span class="sxs-lookup"><span data-stu-id="04f14-p101">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="04f14-143">型</span><span class="sxs-lookup"><span data-stu-id="04f14-143">Type</span></span>

*   <span data-ttu-id="04f14-144">String</span><span class="sxs-lookup"><span data-stu-id="04f14-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="04f14-145">要件</span><span class="sxs-lookup"><span data-stu-id="04f14-145">Requirements</span></span>

|<span data-ttu-id="04f14-146">要件</span><span class="sxs-lookup"><span data-stu-id="04f14-146">Requirement</span></span>| <span data-ttu-id="04f14-147">値</span><span class="sxs-lookup"><span data-stu-id="04f14-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="04f14-148">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="04f14-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04f14-149">1.0</span><span class="sxs-lookup"><span data-stu-id="04f14-149">1.0</span></span>|
|[<span data-ttu-id="04f14-150">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="04f14-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04f14-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04f14-151">ReadItem</span></span>|
|[<span data-ttu-id="04f14-152">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="04f14-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04f14-153">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="04f14-153">Compose or Read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="04f14-154">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="04f14-154">OWAView :String</span></span>

<span data-ttu-id="04f14-155">Outlook Web App の現在のビューを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="04f14-155">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="04f14-156">返される文字列は、値 `OneColumn`、`TwoColumns`、または `ThreeColumns` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="04f14-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="04f14-157">ホスト アプリケーションが Outlook Web App ではない場合、このプロパティにアクセスすると `undefined` が返されます。</span><span class="sxs-lookup"><span data-stu-id="04f14-157">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="04f14-158">Outlook Web App には、画面とウィンドウの幅、および表示可能な列数に応じて 3 つのビューがあります。</span><span class="sxs-lookup"><span data-stu-id="04f14-158">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="04f14-p102">画面幅が狭い場合に表示される `OneColumn`。Outlook Web App は、この単一列レイアウトを使用してスマートフォンの画面全体への表示を行います。</span><span class="sxs-lookup"><span data-stu-id="04f14-p102">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="04f14-p103">画面幅がやや広い場合に表示される `TwoColumns`。Outlook Web App は、ほとんどのタブレットでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="04f14-p103">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="04f14-p104">画面幅が広い場合に表示される `ThreeColumns`。Outlook Web App は、デスクトップ コンピューターのフル スクリーン ウィンドウなどでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="04f14-p104">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="04f14-165">型</span><span class="sxs-lookup"><span data-stu-id="04f14-165">Type</span></span>

*   <span data-ttu-id="04f14-166">String</span><span class="sxs-lookup"><span data-stu-id="04f14-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="04f14-167">要件</span><span class="sxs-lookup"><span data-stu-id="04f14-167">Requirements</span></span>

|<span data-ttu-id="04f14-168">要件</span><span class="sxs-lookup"><span data-stu-id="04f14-168">Requirement</span></span>| <span data-ttu-id="04f14-169">値</span><span class="sxs-lookup"><span data-stu-id="04f14-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="04f14-170">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="04f14-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04f14-171">1.0</span><span class="sxs-lookup"><span data-stu-id="04f14-171">1.0</span></span>|
|[<span data-ttu-id="04f14-172">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="04f14-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04f14-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="04f14-173">ReadItem</span></span>|
|[<span data-ttu-id="04f14-174">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="04f14-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04f14-175">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="04f14-175">Compose or Read</span></span>|
