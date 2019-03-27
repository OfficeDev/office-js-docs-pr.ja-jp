---
title: Office.--の要件セット1.1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 8ce3242a9e007e9aa686ddd098ab23ee922276e4
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871193"
---
# <a name="diagnostics"></a><span data-ttu-id="be1ad-102">診断</span><span class="sxs-lookup"><span data-stu-id="be1ad-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="be1ad-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="be1ad-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="be1ad-104">Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="be1ad-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="be1ad-105">要件</span><span class="sxs-lookup"><span data-stu-id="be1ad-105">Requirements</span></span>

|<span data-ttu-id="be1ad-106">要件</span><span class="sxs-lookup"><span data-stu-id="be1ad-106">Requirement</span></span>| <span data-ttu-id="be1ad-107">値</span><span class="sxs-lookup"><span data-stu-id="be1ad-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="be1ad-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="be1ad-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be1ad-109">1.0</span><span class="sxs-lookup"><span data-stu-id="be1ad-109">1.0</span></span>|
|[<span data-ttu-id="be1ad-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="be1ad-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be1ad-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be1ad-111">ReadItem</span></span>|
|[<span data-ttu-id="be1ad-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="be1ad-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be1ad-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="be1ad-113">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="be1ad-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="be1ad-114">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="be1ad-115">hostName :String</span><span class="sxs-lookup"><span data-stu-id="be1ad-115">hostName :String</span></span>

<span data-ttu-id="be1ad-116">ホスト アプリケーションの名前を表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="be1ad-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="be1ad-117">文字列は、値 `Outlook`、`OutlookIOS`、`OutlookWebApp` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="be1ad-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="be1ad-118">型</span><span class="sxs-lookup"><span data-stu-id="be1ad-118">Type</span></span>

*   <span data-ttu-id="be1ad-119">String</span><span class="sxs-lookup"><span data-stu-id="be1ad-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="be1ad-120">要件</span><span class="sxs-lookup"><span data-stu-id="be1ad-120">Requirements</span></span>

|<span data-ttu-id="be1ad-121">要件</span><span class="sxs-lookup"><span data-stu-id="be1ad-121">Requirement</span></span>| <span data-ttu-id="be1ad-122">値</span><span class="sxs-lookup"><span data-stu-id="be1ad-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="be1ad-123">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="be1ad-123">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be1ad-124">1.0</span><span class="sxs-lookup"><span data-stu-id="be1ad-124">1.0</span></span>|
|[<span data-ttu-id="be1ad-125">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="be1ad-125">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be1ad-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be1ad-126">ReadItem</span></span>|
|[<span data-ttu-id="be1ad-127">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="be1ad-127">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be1ad-128">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="be1ad-128">Compose or Read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="be1ad-129">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="be1ad-129">hostVersion :String</span></span>

<span data-ttu-id="be1ad-130">ホスト アプリケーションまたは Exchange Server のバージョンを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="be1ad-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="be1ad-p101">メール アドインを Outlook デスクトップ クライアントまたは Outlook for iOS で実行している場合、`hostVersion` プロパティは、ホスト アプリケーションである Outlook のバージョンを返します。Outlook Web App では、プロパティは、Exchange Server のバージョンを返します。たとえば、文字列 `15.0.468.0` です。</span><span class="sxs-lookup"><span data-stu-id="be1ad-p101">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="be1ad-134">型</span><span class="sxs-lookup"><span data-stu-id="be1ad-134">Type</span></span>

*   <span data-ttu-id="be1ad-135">String</span><span class="sxs-lookup"><span data-stu-id="be1ad-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="be1ad-136">要件</span><span class="sxs-lookup"><span data-stu-id="be1ad-136">Requirements</span></span>

|<span data-ttu-id="be1ad-137">要件</span><span class="sxs-lookup"><span data-stu-id="be1ad-137">Requirement</span></span>| <span data-ttu-id="be1ad-138">値</span><span class="sxs-lookup"><span data-stu-id="be1ad-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="be1ad-139">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="be1ad-139">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be1ad-140">1.0</span><span class="sxs-lookup"><span data-stu-id="be1ad-140">1.0</span></span>|
|[<span data-ttu-id="be1ad-141">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="be1ad-141">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be1ad-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be1ad-142">ReadItem</span></span>|
|[<span data-ttu-id="be1ad-143">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="be1ad-143">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be1ad-144">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="be1ad-144">Compose or Read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="be1ad-145">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="be1ad-145">OWAView :String</span></span>

<span data-ttu-id="be1ad-146">Outlook Web App の現在のビューを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="be1ad-146">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="be1ad-147">返される文字列は、値 `OneColumn`、`TwoColumns`、または `ThreeColumns` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="be1ad-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="be1ad-148">ホスト アプリケーションが Outlook Web App ではない場合、このプロパティにアクセスすると `undefined` が返されます。</span><span class="sxs-lookup"><span data-stu-id="be1ad-148">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="be1ad-149">Outlook Web App には、画面とウィンドウの幅、および表示可能な列数に応じて 3 つのビューがあります。</span><span class="sxs-lookup"><span data-stu-id="be1ad-149">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="be1ad-p102">画面幅が狭い場合に表示される `OneColumn`。Outlook Web App は、この単一列レイアウトを使用してスマートフォンの画面全体への表示を行います。</span><span class="sxs-lookup"><span data-stu-id="be1ad-p102">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="be1ad-p103">画面幅がやや広い場合に表示される `TwoColumns`。Outlook Web App は、ほとんどのタブレットでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="be1ad-p103">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="be1ad-p104">画面幅が広い場合に表示される `ThreeColumns`。Outlook Web App は、デスクトップ コンピューターのフル スクリーン ウィンドウなどでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="be1ad-p104">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="be1ad-156">型</span><span class="sxs-lookup"><span data-stu-id="be1ad-156">Type</span></span>

*   <span data-ttu-id="be1ad-157">String</span><span class="sxs-lookup"><span data-stu-id="be1ad-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="be1ad-158">要件</span><span class="sxs-lookup"><span data-stu-id="be1ad-158">Requirements</span></span>

|<span data-ttu-id="be1ad-159">要件</span><span class="sxs-lookup"><span data-stu-id="be1ad-159">Requirement</span></span>| <span data-ttu-id="be1ad-160">値</span><span class="sxs-lookup"><span data-stu-id="be1ad-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="be1ad-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="be1ad-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be1ad-162">1.0</span><span class="sxs-lookup"><span data-stu-id="be1ad-162">1.0</span></span>|
|[<span data-ttu-id="be1ad-163">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="be1ad-163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be1ad-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be1ad-164">ReadItem</span></span>|
|[<span data-ttu-id="be1ad-165">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="be1ad-165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be1ad-166">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="be1ad-166">Compose or Read</span></span>|
