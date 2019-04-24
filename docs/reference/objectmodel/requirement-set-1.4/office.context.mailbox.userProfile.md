---
title: Office.-mailbox-要件セット1.4
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 2798b07b3353e9d89f757a22e6bed19dbd94a1c5
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450318"
---
# <a name="userprofile"></a><span data-ttu-id="c09d5-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="c09d5-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="c09d5-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="c09d5-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="c09d5-104">要件</span><span class="sxs-lookup"><span data-stu-id="c09d5-104">Requirements</span></span>

|<span data-ttu-id="c09d5-105">要件</span><span class="sxs-lookup"><span data-stu-id="c09d5-105">Requirement</span></span>| <span data-ttu-id="c09d5-106">値</span><span class="sxs-lookup"><span data-stu-id="c09d5-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="c09d5-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c09d5-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c09d5-108">1.0</span><span class="sxs-lookup"><span data-stu-id="c09d5-108">1.0</span></span>|
|[<span data-ttu-id="c09d5-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c09d5-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c09d5-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c09d5-110">ReadItem</span></span>|
|[<span data-ttu-id="c09d5-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c09d5-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c09d5-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c09d5-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="c09d5-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="c09d5-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="c09d5-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="c09d5-114">displayName :String</span></span>

<span data-ttu-id="c09d5-115">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="c09d5-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="c09d5-116">型</span><span class="sxs-lookup"><span data-stu-id="c09d5-116">Type</span></span>

*   <span data-ttu-id="c09d5-117">String</span><span class="sxs-lookup"><span data-stu-id="c09d5-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c09d5-118">要件</span><span class="sxs-lookup"><span data-stu-id="c09d5-118">Requirements</span></span>

|<span data-ttu-id="c09d5-119">要件</span><span class="sxs-lookup"><span data-stu-id="c09d5-119">Requirement</span></span>| <span data-ttu-id="c09d5-120">値</span><span class="sxs-lookup"><span data-stu-id="c09d5-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="c09d5-121">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c09d5-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c09d5-122">1.0</span><span class="sxs-lookup"><span data-stu-id="c09d5-122">1.0</span></span>|
|[<span data-ttu-id="c09d5-123">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c09d5-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c09d5-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c09d5-124">ReadItem</span></span>|
|[<span data-ttu-id="c09d5-125">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c09d5-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c09d5-126">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c09d5-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c09d5-127">例</span><span class="sxs-lookup"><span data-stu-id="c09d5-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="c09d5-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="c09d5-128">emailAddress :String</span></span>

<span data-ttu-id="c09d5-129">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="c09d5-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="c09d5-130">型</span><span class="sxs-lookup"><span data-stu-id="c09d5-130">Type</span></span>

*   <span data-ttu-id="c09d5-131">String</span><span class="sxs-lookup"><span data-stu-id="c09d5-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c09d5-132">要件</span><span class="sxs-lookup"><span data-stu-id="c09d5-132">Requirements</span></span>

|<span data-ttu-id="c09d5-133">要件</span><span class="sxs-lookup"><span data-stu-id="c09d5-133">Requirement</span></span>| <span data-ttu-id="c09d5-134">値</span><span class="sxs-lookup"><span data-stu-id="c09d5-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="c09d5-135">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c09d5-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c09d5-136">1.0</span><span class="sxs-lookup"><span data-stu-id="c09d5-136">1.0</span></span>|
|[<span data-ttu-id="c09d5-137">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c09d5-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c09d5-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c09d5-138">ReadItem</span></span>|
|[<span data-ttu-id="c09d5-139">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c09d5-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c09d5-140">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c09d5-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c09d5-141">例</span><span class="sxs-lookup"><span data-stu-id="c09d5-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="c09d5-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="c09d5-142">timeZone :String</span></span>

<span data-ttu-id="c09d5-143">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="c09d5-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="c09d5-144">型</span><span class="sxs-lookup"><span data-stu-id="c09d5-144">Type</span></span>

*   <span data-ttu-id="c09d5-145">String</span><span class="sxs-lookup"><span data-stu-id="c09d5-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c09d5-146">要件</span><span class="sxs-lookup"><span data-stu-id="c09d5-146">Requirements</span></span>

|<span data-ttu-id="c09d5-147">要件</span><span class="sxs-lookup"><span data-stu-id="c09d5-147">Requirement</span></span>| <span data-ttu-id="c09d5-148">値</span><span class="sxs-lookup"><span data-stu-id="c09d5-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="c09d5-149">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c09d5-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c09d5-150">1.0</span><span class="sxs-lookup"><span data-stu-id="c09d5-150">1.0</span></span>|
|[<span data-ttu-id="c09d5-151">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c09d5-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c09d5-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c09d5-152">ReadItem</span></span>|
|[<span data-ttu-id="c09d5-153">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c09d5-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c09d5-154">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c09d5-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c09d5-155">例</span><span class="sxs-lookup"><span data-stu-id="c09d5-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
