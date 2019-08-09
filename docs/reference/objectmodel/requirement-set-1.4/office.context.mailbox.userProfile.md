---
title: Office.-mailbox-要件セット1.4
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 7a728ebbec0136e0b2eddfb4402e45abe3f02ad4
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268636"
---
# <a name="userprofile"></a><span data-ttu-id="e8ee3-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="e8ee3-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="e8ee3-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="e8ee3-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="e8ee3-104">要件</span><span class="sxs-lookup"><span data-stu-id="e8ee3-104">Requirements</span></span>

|<span data-ttu-id="e8ee3-105">要件</span><span class="sxs-lookup"><span data-stu-id="e8ee3-105">Requirement</span></span>| <span data-ttu-id="e8ee3-106">値</span><span class="sxs-lookup"><span data-stu-id="e8ee3-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="e8ee3-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e8ee3-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e8ee3-108">1.0</span><span class="sxs-lookup"><span data-stu-id="e8ee3-108">1.0</span></span>|
|[<span data-ttu-id="e8ee3-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e8ee3-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e8ee3-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e8ee3-110">ReadItem</span></span>|
|[<span data-ttu-id="e8ee3-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e8ee3-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e8ee3-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e8ee3-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e8ee3-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="e8ee3-113">Members and methods</span></span>

| <span data-ttu-id="e8ee3-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="e8ee3-114">Member</span></span> | <span data-ttu-id="e8ee3-115">種類</span><span class="sxs-lookup"><span data-stu-id="e8ee3-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e8ee3-116">displayName</span><span class="sxs-lookup"><span data-stu-id="e8ee3-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="e8ee3-117">Member</span><span class="sxs-lookup"><span data-stu-id="e8ee3-117">Member</span></span> |
| [<span data-ttu-id="e8ee3-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="e8ee3-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="e8ee3-119">Member</span><span class="sxs-lookup"><span data-stu-id="e8ee3-119">Member</span></span> |
| [<span data-ttu-id="e8ee3-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="e8ee3-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="e8ee3-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="e8ee3-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="e8ee3-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="e8ee3-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="e8ee3-123">displayName: String</span><span class="sxs-lookup"><span data-stu-id="e8ee3-123">displayName: String</span></span>

<span data-ttu-id="e8ee3-124">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="e8ee3-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="e8ee3-125">型</span><span class="sxs-lookup"><span data-stu-id="e8ee3-125">Type</span></span>

*   <span data-ttu-id="e8ee3-126">String</span><span class="sxs-lookup"><span data-stu-id="e8ee3-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e8ee3-127">要件</span><span class="sxs-lookup"><span data-stu-id="e8ee3-127">Requirements</span></span>

|<span data-ttu-id="e8ee3-128">要件</span><span class="sxs-lookup"><span data-stu-id="e8ee3-128">Requirement</span></span>| <span data-ttu-id="e8ee3-129">値</span><span class="sxs-lookup"><span data-stu-id="e8ee3-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="e8ee3-130">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e8ee3-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e8ee3-131">1.0</span><span class="sxs-lookup"><span data-stu-id="e8ee3-131">1.0</span></span>|
|[<span data-ttu-id="e8ee3-132">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e8ee3-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e8ee3-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e8ee3-133">ReadItem</span></span>|
|[<span data-ttu-id="e8ee3-134">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e8ee3-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e8ee3-135">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e8ee3-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e8ee3-136">例</span><span class="sxs-lookup"><span data-stu-id="e8ee3-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

#### <a name="emailaddress-string"></a><span data-ttu-id="e8ee3-137">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="e8ee3-137">emailAddress: String</span></span>

<span data-ttu-id="e8ee3-138">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="e8ee3-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="e8ee3-139">型</span><span class="sxs-lookup"><span data-stu-id="e8ee3-139">Type</span></span>

*   <span data-ttu-id="e8ee3-140">String</span><span class="sxs-lookup"><span data-stu-id="e8ee3-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e8ee3-141">要件</span><span class="sxs-lookup"><span data-stu-id="e8ee3-141">Requirements</span></span>

|<span data-ttu-id="e8ee3-142">要件</span><span class="sxs-lookup"><span data-stu-id="e8ee3-142">Requirement</span></span>| <span data-ttu-id="e8ee3-143">値</span><span class="sxs-lookup"><span data-stu-id="e8ee3-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="e8ee3-144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e8ee3-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e8ee3-145">1.0</span><span class="sxs-lookup"><span data-stu-id="e8ee3-145">1.0</span></span>|
|[<span data-ttu-id="e8ee3-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e8ee3-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e8ee3-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e8ee3-147">ReadItem</span></span>|
|[<span data-ttu-id="e8ee3-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e8ee3-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e8ee3-149">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e8ee3-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e8ee3-150">例</span><span class="sxs-lookup"><span data-stu-id="e8ee3-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

#### <a name="timezone-string"></a><span data-ttu-id="e8ee3-151">timeZone: String</span><span class="sxs-lookup"><span data-stu-id="e8ee3-151">timeZone: String</span></span>

<span data-ttu-id="e8ee3-152">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="e8ee3-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="e8ee3-153">型</span><span class="sxs-lookup"><span data-stu-id="e8ee3-153">Type</span></span>

*   <span data-ttu-id="e8ee3-154">String</span><span class="sxs-lookup"><span data-stu-id="e8ee3-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e8ee3-155">要件</span><span class="sxs-lookup"><span data-stu-id="e8ee3-155">Requirements</span></span>

|<span data-ttu-id="e8ee3-156">要件</span><span class="sxs-lookup"><span data-stu-id="e8ee3-156">Requirement</span></span>| <span data-ttu-id="e8ee3-157">値</span><span class="sxs-lookup"><span data-stu-id="e8ee3-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="e8ee3-158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e8ee3-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e8ee3-159">1.0</span><span class="sxs-lookup"><span data-stu-id="e8ee3-159">1.0</span></span>|
|[<span data-ttu-id="e8ee3-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e8ee3-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e8ee3-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e8ee3-161">ReadItem</span></span>|
|[<span data-ttu-id="e8ee3-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e8ee3-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e8ee3-163">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e8ee3-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e8ee3-164">例</span><span class="sxs-lookup"><span data-stu-id="e8ee3-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
