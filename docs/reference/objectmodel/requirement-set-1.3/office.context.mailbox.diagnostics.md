---
title: Office.context.mailbox.diagnostics - 要件セット 1.3
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: a0edb17bebed0965b6fc642f386b94aaca2e4ad3
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433132"
---
# <a name="diagnostics"></a><span data-ttu-id="dcb9c-102">診断</span><span class="sxs-lookup"><span data-stu-id="dcb9c-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="dcb9c-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="dcb9c-103">Office.context.mailbox.diagnostics</span></span>

<span data-ttu-id="dcb9c-104">Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="dcb9c-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="dcb9c-105">要件</span><span class="sxs-lookup"><span data-stu-id="dcb9c-105">Requirements</span></span>

|<span data-ttu-id="dcb9c-106">要件</span><span class="sxs-lookup"><span data-stu-id="dcb9c-106">Requirement</span></span>| <span data-ttu-id="dcb9c-107">値</span><span class="sxs-lookup"><span data-stu-id="dcb9c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="dcb9c-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dcb9c-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dcb9c-109">1.0</span><span class="sxs-lookup"><span data-stu-id="dcb9c-109">1.0</span></span>|
|[<span data-ttu-id="dcb9c-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dcb9c-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dcb9c-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dcb9c-111">ReadItem</span></span>|
|[<span data-ttu-id="dcb9c-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dcb9c-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dcb9c-113">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="dcb9c-113">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="dcb9c-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="dcb9c-114">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="dcb9c-115">hostName :String</span><span class="sxs-lookup"><span data-stu-id="dcb9c-115">hostName :String</span></span>

<span data-ttu-id="dcb9c-116">ホスト アプリケーションの名前を表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="dcb9c-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="dcb9c-117">文字列は、値 `Outlook`、`OutlookIOS`、`OutlookWebApp` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="dcb9c-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, , or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="dcb9c-118">型:</span><span class="sxs-lookup"><span data-stu-id="dcb9c-118">Type:</span></span>

*   <span data-ttu-id="dcb9c-119">String</span><span class="sxs-lookup"><span data-stu-id="dcb9c-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dcb9c-120">要件</span><span class="sxs-lookup"><span data-stu-id="dcb9c-120">Requirements</span></span>

|<span data-ttu-id="dcb9c-121">要件</span><span class="sxs-lookup"><span data-stu-id="dcb9c-121">Requirement</span></span>| <span data-ttu-id="dcb9c-122">値</span><span class="sxs-lookup"><span data-stu-id="dcb9c-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="dcb9c-123">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dcb9c-123">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dcb9c-124">1.0</span><span class="sxs-lookup"><span data-stu-id="dcb9c-124">1.0</span></span>|
|[<span data-ttu-id="dcb9c-125">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dcb9c-125">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dcb9c-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dcb9c-126">ReadItem</span></span>|
|[<span data-ttu-id="dcb9c-127">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dcb9c-127">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dcb9c-128">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="dcb9c-128">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="dcb9c-129">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="dcb9c-129">hostVersion :String</span></span>

<span data-ttu-id="dcb9c-130">ホスト アプリケーションまたは Exchange Server のバージョンを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="dcb9c-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="dcb9c-p101">メール アドインを Outlook デスクトップ クライアントまたは Outlook for iOS で実行している場合、`hostVersion` プロパティは、ホスト アプリケーションである Outlook のバージョンを返します。Outlook Web App では、プロパティは、Exchange Server のバージョンを返します。たとえば、文字列 `15.0.468.0` です。</span><span class="sxs-lookup"><span data-stu-id="dcb9c-p101">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="dcb9c-134">種類:</span><span class="sxs-lookup"><span data-stu-id="dcb9c-134">Type:</span></span>

*   <span data-ttu-id="dcb9c-135">String</span><span class="sxs-lookup"><span data-stu-id="dcb9c-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dcb9c-136">要件</span><span class="sxs-lookup"><span data-stu-id="dcb9c-136">Requirements</span></span>

|<span data-ttu-id="dcb9c-137">要件</span><span class="sxs-lookup"><span data-stu-id="dcb9c-137">Requirement</span></span>| <span data-ttu-id="dcb9c-138">値</span><span class="sxs-lookup"><span data-stu-id="dcb9c-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="dcb9c-139">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dcb9c-139">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dcb9c-140">1.0</span><span class="sxs-lookup"><span data-stu-id="dcb9c-140">1.0</span></span>|
|[<span data-ttu-id="dcb9c-141">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dcb9c-141">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dcb9c-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dcb9c-142">ReadItem</span></span>|
|[<span data-ttu-id="dcb9c-143">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dcb9c-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dcb9c-144">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="dcb9c-144">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="dcb9c-145">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="dcb9c-145">OWAView :String</span></span>

<span data-ttu-id="dcb9c-146">Outlook Web App の現在のビューを表す文字列を取得します。</span><span class="sxs-lookup"><span data-stu-id="dcb9c-146">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="dcb9c-147">返される文字列は、値 `OneColumn`、`TwoColumns`、または `ThreeColumns` のいずれかになります。</span><span class="sxs-lookup"><span data-stu-id="dcb9c-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="dcb9c-148">ホスト アプリケーションが Outlook Web App ではない場合、このプロパティにアクセスすると `undefined` が返されます。</span><span class="sxs-lookup"><span data-stu-id="dcb9c-148">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="dcb9c-149">Outlook Web App には、画面とウィンドウの幅、および表示可能な列数に応じて 3 つのビューがあります。</span><span class="sxs-lookup"><span data-stu-id="dcb9c-149">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="dcb9c-p102">画面幅が狭い場合に表示される `OneColumn`。Outlook Web App は、この単一列レイアウトを使用してスマートフォンの画面全体への表示を行います。</span><span class="sxs-lookup"><span data-stu-id="dcb9c-p102">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="dcb9c-p103">画面幅がやや広い場合に表示される `TwoColumns`。Outlook Web App は、ほとんどのタブレットでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="dcb9c-p103">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="dcb9c-p104">画面幅が広い場合に表示される `ThreeColumns`。Outlook Web App は、デスクトップ コンピューターのフル スクリーン ウィンドウなどでこのビューを使用します。</span><span class="sxs-lookup"><span data-stu-id="dcb9c-p104">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="dcb9c-156">型:</span><span class="sxs-lookup"><span data-stu-id="dcb9c-156">Type:</span></span>

*   <span data-ttu-id="dcb9c-157">String</span><span class="sxs-lookup"><span data-stu-id="dcb9c-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dcb9c-158">要件</span><span class="sxs-lookup"><span data-stu-id="dcb9c-158">Requirements</span></span>

|<span data-ttu-id="dcb9c-159">要件</span><span class="sxs-lookup"><span data-stu-id="dcb9c-159">Requirement</span></span>| <span data-ttu-id="dcb9c-160">値</span><span class="sxs-lookup"><span data-stu-id="dcb9c-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="dcb9c-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dcb9c-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dcb9c-162">1.0</span><span class="sxs-lookup"><span data-stu-id="dcb9c-162">1.0</span></span>|
|[<span data-ttu-id="dcb9c-163">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dcb9c-163">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dcb9c-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dcb9c-164">ReadItem</span></span>|
|[<span data-ttu-id="dcb9c-165">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dcb9c-165">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dcb9c-166">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="dcb9c-166">Compose or read</span></span>|