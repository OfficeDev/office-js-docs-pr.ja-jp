---
title: Office.-mailbox-要件セット1.5
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: fc20497cc8df8d091ba0195f7dca9b283ff4d1c2
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871018"
---
# <a name="userprofile"></a><span data-ttu-id="15503-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="15503-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="15503-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="15503-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="15503-104">要件</span><span class="sxs-lookup"><span data-stu-id="15503-104">Requirements</span></span>

|<span data-ttu-id="15503-105">要件</span><span class="sxs-lookup"><span data-stu-id="15503-105">Requirement</span></span>| <span data-ttu-id="15503-106">値</span><span class="sxs-lookup"><span data-stu-id="15503-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="15503-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="15503-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="15503-108">1.0</span><span class="sxs-lookup"><span data-stu-id="15503-108">1.0</span></span>|
|[<span data-ttu-id="15503-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="15503-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="15503-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="15503-110">ReadItem</span></span>|
|[<span data-ttu-id="15503-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="15503-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="15503-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="15503-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="15503-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="15503-113">Members and methods</span></span>

| <span data-ttu-id="15503-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="15503-114">Member</span></span> | <span data-ttu-id="15503-115">種類</span><span class="sxs-lookup"><span data-stu-id="15503-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="15503-116">displayName</span><span class="sxs-lookup"><span data-stu-id="15503-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="15503-117">Member</span><span class="sxs-lookup"><span data-stu-id="15503-117">Member</span></span> |
| [<span data-ttu-id="15503-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="15503-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="15503-119">Member</span><span class="sxs-lookup"><span data-stu-id="15503-119">Member</span></span> |
| [<span data-ttu-id="15503-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="15503-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="15503-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="15503-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="15503-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="15503-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="15503-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="15503-123">displayName :String</span></span>

<span data-ttu-id="15503-124">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="15503-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="15503-125">型</span><span class="sxs-lookup"><span data-stu-id="15503-125">Type</span></span>

*   <span data-ttu-id="15503-126">String</span><span class="sxs-lookup"><span data-stu-id="15503-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="15503-127">要件</span><span class="sxs-lookup"><span data-stu-id="15503-127">Requirements</span></span>

|<span data-ttu-id="15503-128">要件</span><span class="sxs-lookup"><span data-stu-id="15503-128">Requirement</span></span>| <span data-ttu-id="15503-129">値</span><span class="sxs-lookup"><span data-stu-id="15503-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="15503-130">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="15503-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="15503-131">1.0</span><span class="sxs-lookup"><span data-stu-id="15503-131">1.0</span></span>|
|[<span data-ttu-id="15503-132">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="15503-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="15503-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="15503-133">ReadItem</span></span>|
|[<span data-ttu-id="15503-134">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="15503-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="15503-135">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="15503-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="15503-136">例</span><span class="sxs-lookup"><span data-stu-id="15503-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="15503-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="15503-137">emailAddress :String</span></span>

<span data-ttu-id="15503-138">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="15503-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="15503-139">型</span><span class="sxs-lookup"><span data-stu-id="15503-139">Type</span></span>

*   <span data-ttu-id="15503-140">String</span><span class="sxs-lookup"><span data-stu-id="15503-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="15503-141">要件</span><span class="sxs-lookup"><span data-stu-id="15503-141">Requirements</span></span>

|<span data-ttu-id="15503-142">要件</span><span class="sxs-lookup"><span data-stu-id="15503-142">Requirement</span></span>| <span data-ttu-id="15503-143">値</span><span class="sxs-lookup"><span data-stu-id="15503-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="15503-144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="15503-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="15503-145">1.0</span><span class="sxs-lookup"><span data-stu-id="15503-145">1.0</span></span>|
|[<span data-ttu-id="15503-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="15503-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="15503-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="15503-147">ReadItem</span></span>|
|[<span data-ttu-id="15503-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="15503-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="15503-149">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="15503-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="15503-150">例</span><span class="sxs-lookup"><span data-stu-id="15503-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="15503-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="15503-151">timeZone :String</span></span>

<span data-ttu-id="15503-152">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="15503-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="15503-153">型</span><span class="sxs-lookup"><span data-stu-id="15503-153">Type</span></span>

*   <span data-ttu-id="15503-154">String</span><span class="sxs-lookup"><span data-stu-id="15503-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="15503-155">要件</span><span class="sxs-lookup"><span data-stu-id="15503-155">Requirements</span></span>

|<span data-ttu-id="15503-156">要件</span><span class="sxs-lookup"><span data-stu-id="15503-156">Requirement</span></span>| <span data-ttu-id="15503-157">値</span><span class="sxs-lookup"><span data-stu-id="15503-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="15503-158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="15503-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="15503-159">1.0</span><span class="sxs-lookup"><span data-stu-id="15503-159">1.0</span></span>|
|[<span data-ttu-id="15503-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="15503-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="15503-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="15503-161">ReadItem</span></span>|
|[<span data-ttu-id="15503-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="15503-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="15503-163">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="15503-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="15503-164">例</span><span class="sxs-lookup"><span data-stu-id="15503-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
