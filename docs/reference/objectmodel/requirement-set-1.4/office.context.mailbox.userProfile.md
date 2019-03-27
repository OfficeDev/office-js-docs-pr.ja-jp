---
title: Office.-mailbox-要件セット1.4
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 2798b07b3353e9d89f757a22e6bed19dbd94a1c5
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870045"
---
# <a name="userprofile"></a><span data-ttu-id="2d8fa-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="2d8fa-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="2d8fa-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="2d8fa-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="2d8fa-104">要件</span><span class="sxs-lookup"><span data-stu-id="2d8fa-104">Requirements</span></span>

|<span data-ttu-id="2d8fa-105">要件</span><span class="sxs-lookup"><span data-stu-id="2d8fa-105">Requirement</span></span>| <span data-ttu-id="2d8fa-106">値</span><span class="sxs-lookup"><span data-stu-id="2d8fa-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d8fa-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2d8fa-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d8fa-108">1.0</span><span class="sxs-lookup"><span data-stu-id="2d8fa-108">1.0</span></span>|
|[<span data-ttu-id="2d8fa-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2d8fa-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d8fa-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d8fa-110">ReadItem</span></span>|
|[<span data-ttu-id="2d8fa-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2d8fa-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d8fa-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2d8fa-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="2d8fa-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="2d8fa-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="2d8fa-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="2d8fa-114">displayName :String</span></span>

<span data-ttu-id="2d8fa-115">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="2d8fa-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="2d8fa-116">型</span><span class="sxs-lookup"><span data-stu-id="2d8fa-116">Type</span></span>

*   <span data-ttu-id="2d8fa-117">String</span><span class="sxs-lookup"><span data-stu-id="2d8fa-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2d8fa-118">要件</span><span class="sxs-lookup"><span data-stu-id="2d8fa-118">Requirements</span></span>

|<span data-ttu-id="2d8fa-119">要件</span><span class="sxs-lookup"><span data-stu-id="2d8fa-119">Requirement</span></span>| <span data-ttu-id="2d8fa-120">値</span><span class="sxs-lookup"><span data-stu-id="2d8fa-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d8fa-121">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2d8fa-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d8fa-122">1.0</span><span class="sxs-lookup"><span data-stu-id="2d8fa-122">1.0</span></span>|
|[<span data-ttu-id="2d8fa-123">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2d8fa-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d8fa-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d8fa-124">ReadItem</span></span>|
|[<span data-ttu-id="2d8fa-125">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2d8fa-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d8fa-126">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2d8fa-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2d8fa-127">例</span><span class="sxs-lookup"><span data-stu-id="2d8fa-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="2d8fa-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="2d8fa-128">emailAddress :String</span></span>

<span data-ttu-id="2d8fa-129">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="2d8fa-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="2d8fa-130">型</span><span class="sxs-lookup"><span data-stu-id="2d8fa-130">Type</span></span>

*   <span data-ttu-id="2d8fa-131">String</span><span class="sxs-lookup"><span data-stu-id="2d8fa-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2d8fa-132">要件</span><span class="sxs-lookup"><span data-stu-id="2d8fa-132">Requirements</span></span>

|<span data-ttu-id="2d8fa-133">要件</span><span class="sxs-lookup"><span data-stu-id="2d8fa-133">Requirement</span></span>| <span data-ttu-id="2d8fa-134">値</span><span class="sxs-lookup"><span data-stu-id="2d8fa-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d8fa-135">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2d8fa-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d8fa-136">1.0</span><span class="sxs-lookup"><span data-stu-id="2d8fa-136">1.0</span></span>|
|[<span data-ttu-id="2d8fa-137">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2d8fa-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d8fa-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d8fa-138">ReadItem</span></span>|
|[<span data-ttu-id="2d8fa-139">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2d8fa-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d8fa-140">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2d8fa-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2d8fa-141">例</span><span class="sxs-lookup"><span data-stu-id="2d8fa-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="2d8fa-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="2d8fa-142">timeZone :String</span></span>

<span data-ttu-id="2d8fa-143">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="2d8fa-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="2d8fa-144">型</span><span class="sxs-lookup"><span data-stu-id="2d8fa-144">Type</span></span>

*   <span data-ttu-id="2d8fa-145">String</span><span class="sxs-lookup"><span data-stu-id="2d8fa-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2d8fa-146">要件</span><span class="sxs-lookup"><span data-stu-id="2d8fa-146">Requirements</span></span>

|<span data-ttu-id="2d8fa-147">要件</span><span class="sxs-lookup"><span data-stu-id="2d8fa-147">Requirement</span></span>| <span data-ttu-id="2d8fa-148">値</span><span class="sxs-lookup"><span data-stu-id="2d8fa-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d8fa-149">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2d8fa-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d8fa-150">1.0</span><span class="sxs-lookup"><span data-stu-id="2d8fa-150">1.0</span></span>|
|[<span data-ttu-id="2d8fa-151">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2d8fa-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d8fa-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2d8fa-152">ReadItem</span></span>|
|[<span data-ttu-id="2d8fa-153">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2d8fa-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d8fa-154">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2d8fa-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2d8fa-155">例</span><span class="sxs-lookup"><span data-stu-id="2d8fa-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
