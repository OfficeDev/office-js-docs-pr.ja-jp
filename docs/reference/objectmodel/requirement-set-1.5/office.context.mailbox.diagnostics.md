---
title: Office.context.mailbox.diagnostics - 要件セット 1.5
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 3a00c714a766bf13c83a63fc30a564a88f421a09
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067945"
---
# <a name="diagnostics"></a><span data-ttu-id="e4efd-102">診断</span><span class="sxs-lookup"><span data-stu-id="e4efd-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="e4efd-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="e4efd-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="e4efd-104">Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="e4efd-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e4efd-105">要件</span><span class="sxs-lookup"><span data-stu-id="e4efd-105">Requirements</span></span>

|<span data-ttu-id="e4efd-106">要件</span><span class="sxs-lookup"><span data-stu-id="e4efd-106">Requirement</span></span>| <span data-ttu-id="e4efd-107">値</span><span class="sxs-lookup"><span data-stu-id="e4efd-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e4efd-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e4efd-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e4efd-109">1.0</span><span class="sxs-lookup"><span data-stu-id="e4efd-109">1.0</span></span>|
|[<span data-ttu-id="e4efd-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e4efd-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e4efd-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e4efd-111">ReadItem</span></span>|
|[<span data-ttu-id="e4efd-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e4efd-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e4efd-113">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e4efd-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e4efd-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="e4efd-114">Members and methods</span></span>

| <span data-ttu-id="e4efd-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="e4efd-115">Member</span></span> | <span data-ttu-id="e4efd-116">種類</span><span class="sxs-lookup"><span data-stu-id="e4efd-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e4efd-117">hostName</span><span class="sxs-lookup"><span data-stu-id="e4efd-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="e4efd-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="e4efd-118">Member</span></span> |
| [<span data-ttu-id="e4efd-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="e4efd-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="e4efd-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="e4efd-120">Member</span></span> |
| [<span data-ttu-id="e4efd-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="e4efd-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="e4efd-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="e4efd-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="e4efd-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="e4efd-123">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="e4efd-124">hostName :String</span><span class="sxs-lookup"><span data-stu-id="e4efd-124">hostName :String</span></span>

<span data-ttu-id="e4efd-125">ホスト アプリケーションの名前を表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="e4efd-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="e4efd-126">文字列は、値 `Outlook`、`OutlookIOS`、`OutlookWebApp` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="e4efd-126">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="e4efd-127">Type</span><span class="sxs-lookup"><span data-stu-id="e4efd-127">Type</span></span>

*   <span data-ttu-id="e4efd-128">String</span><span class="sxs-lookup"><span data-stu-id="e4efd-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e4efd-129">要件</span><span class="sxs-lookup"><span data-stu-id="e4efd-129">Requirements</span></span>

|<span data-ttu-id="e4efd-130">要件</span><span class="sxs-lookup"><span data-stu-id="e4efd-130">Requirement</span></span>| <span data-ttu-id="e4efd-131">値</span><span class="sxs-lookup"><span data-stu-id="e4efd-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="e4efd-132">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e4efd-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e4efd-133">1.0</span><span class="sxs-lookup"><span data-stu-id="e4efd-133">1.0</span></span>|
|[<span data-ttu-id="e4efd-134">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e4efd-134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e4efd-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e4efd-135">ReadItem</span></span>|
|[<span data-ttu-id="e4efd-136">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e4efd-136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e4efd-137">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e4efd-137">Compose or Read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="e4efd-138">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="e4efd-138">hostVersion :String</span></span>

<span data-ttu-id="e4efd-139">ホスト アプリケーションまたは Exchange Server のバージョンを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="e4efd-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="e4efd-p101">メール アドインを Outlook デスクトップ クライアントまたは Outlook for iOS で実行している場合、`hostVersion` プロパティは、ホスト アプリケーションである Outlook のバージョンを返します。Outlook Web App では、プロパティは、Exchange Server のバージョンを返します。たとえば、文字列 `15.0.468.0` です。</span><span class="sxs-lookup"><span data-stu-id="e4efd-p101">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="e4efd-143">Type</span><span class="sxs-lookup"><span data-stu-id="e4efd-143">Type</span></span>

*   <span data-ttu-id="e4efd-144">String</span><span class="sxs-lookup"><span data-stu-id="e4efd-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e4efd-145">要件</span><span class="sxs-lookup"><span data-stu-id="e4efd-145">Requirements</span></span>

|<span data-ttu-id="e4efd-146">要件</span><span class="sxs-lookup"><span data-stu-id="e4efd-146">Requirement</span></span>| <span data-ttu-id="e4efd-147">値</span><span class="sxs-lookup"><span data-stu-id="e4efd-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="e4efd-148">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e4efd-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e4efd-149">1.0</span><span class="sxs-lookup"><span data-stu-id="e4efd-149">1.0</span></span>|
|[<span data-ttu-id="e4efd-150">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e4efd-150">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e4efd-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e4efd-151">ReadItem</span></span>|
|[<span data-ttu-id="e4efd-152">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e4efd-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e4efd-153">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e4efd-153">Compose or Read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="e4efd-154">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="e4efd-154">OWAView :String</span></span>

<span data-ttu-id="e4efd-155">Outlook Web App の現在のビューを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="e4efd-155">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="e4efd-156">返される文字列は、値 `OneColumn`、`TwoColumns`、または `ThreeColumns` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="e4efd-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="e4efd-157">ホスト アプリケーションが Outlook Web App ではない場合、このプロパティにアクセスすると `undefined` が返されます。</span><span class="sxs-lookup"><span data-stu-id="e4efd-157">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="e4efd-158">Outlook Web App には、画面とウィンドウの幅、および表示可能な列数に応じて 3 つのビューがあります。</span><span class="sxs-lookup"><span data-stu-id="e4efd-158">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="e4efd-p102">画面幅が狭い場合に表示される `OneColumn`。Outlook Web App は、この単一列レイアウトを使用してスマートフォンの画面全体への表示を行います。</span><span class="sxs-lookup"><span data-stu-id="e4efd-p102">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="e4efd-p103">画面幅がやや広い場合に表示される `TwoColumns`。Outlook Web App は、ほとんどのタブレットでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="e4efd-p103">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="e4efd-p104">画面幅が広い場合に表示される `ThreeColumns`。Outlook Web App は、デスクトップ コンピューターのフル スクリーン ウィンドウなどでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="e4efd-p104">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="e4efd-165">Type</span><span class="sxs-lookup"><span data-stu-id="e4efd-165">Type</span></span>

*   <span data-ttu-id="e4efd-166">String</span><span class="sxs-lookup"><span data-stu-id="e4efd-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e4efd-167">要件</span><span class="sxs-lookup"><span data-stu-id="e4efd-167">Requirements</span></span>

|<span data-ttu-id="e4efd-168">要件</span><span class="sxs-lookup"><span data-stu-id="e4efd-168">Requirement</span></span>| <span data-ttu-id="e4efd-169">値</span><span class="sxs-lookup"><span data-stu-id="e4efd-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="e4efd-170">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e4efd-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e4efd-171">1.0</span><span class="sxs-lookup"><span data-stu-id="e4efd-171">1.0</span></span>|
|[<span data-ttu-id="e4efd-172">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e4efd-172">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e4efd-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e4efd-173">ReadItem</span></span>|
|[<span data-ttu-id="e4efd-174">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e4efd-174">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e4efd-175">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e4efd-175">Compose or Read</span></span>|
