---
title: Office.-mailbox-要件セット1.2
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 8ba2a21b16c51c827155d793241b80c5c510dd5a
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696338"
---
# <a name="userprofile"></a><span data-ttu-id="09874-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="09874-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="09874-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="09874-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="09874-104">要件</span><span class="sxs-lookup"><span data-stu-id="09874-104">Requirements</span></span>

|<span data-ttu-id="09874-105">要件</span><span class="sxs-lookup"><span data-stu-id="09874-105">Requirement</span></span>| <span data-ttu-id="09874-106">値</span><span class="sxs-lookup"><span data-stu-id="09874-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="09874-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="09874-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09874-108">1.0</span><span class="sxs-lookup"><span data-stu-id="09874-108">1.0</span></span>|
|[<span data-ttu-id="09874-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="09874-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09874-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09874-110">ReadItem</span></span>|
|[<span data-ttu-id="09874-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="09874-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="09874-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="09874-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="09874-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="09874-113">Members and methods</span></span>

| <span data-ttu-id="09874-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="09874-114">Member</span></span> | <span data-ttu-id="09874-115">種類</span><span class="sxs-lookup"><span data-stu-id="09874-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="09874-116">displayName</span><span class="sxs-lookup"><span data-stu-id="09874-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="09874-117">Member</span><span class="sxs-lookup"><span data-stu-id="09874-117">Member</span></span> |
| [<span data-ttu-id="09874-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="09874-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="09874-119">Member</span><span class="sxs-lookup"><span data-stu-id="09874-119">Member</span></span> |
| [<span data-ttu-id="09874-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="09874-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="09874-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="09874-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="09874-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="09874-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="09874-123">displayName: String</span><span class="sxs-lookup"><span data-stu-id="09874-123">displayName: String</span></span>

<span data-ttu-id="09874-124">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="09874-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="09874-125">型</span><span class="sxs-lookup"><span data-stu-id="09874-125">Type</span></span>

*   <span data-ttu-id="09874-126">String</span><span class="sxs-lookup"><span data-stu-id="09874-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="09874-127">要件</span><span class="sxs-lookup"><span data-stu-id="09874-127">Requirements</span></span>

|<span data-ttu-id="09874-128">要件</span><span class="sxs-lookup"><span data-stu-id="09874-128">Requirement</span></span>| <span data-ttu-id="09874-129">値</span><span class="sxs-lookup"><span data-stu-id="09874-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="09874-130">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="09874-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09874-131">1.0</span><span class="sxs-lookup"><span data-stu-id="09874-131">1.0</span></span>|
|[<span data-ttu-id="09874-132">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="09874-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09874-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09874-133">ReadItem</span></span>|
|[<span data-ttu-id="09874-134">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="09874-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="09874-135">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="09874-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09874-136">例</span><span class="sxs-lookup"><span data-stu-id="09874-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="09874-137">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="09874-137">emailAddress: String</span></span>

<span data-ttu-id="09874-138">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="09874-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="09874-139">型</span><span class="sxs-lookup"><span data-stu-id="09874-139">Type</span></span>

*   <span data-ttu-id="09874-140">String</span><span class="sxs-lookup"><span data-stu-id="09874-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="09874-141">要件</span><span class="sxs-lookup"><span data-stu-id="09874-141">Requirements</span></span>

|<span data-ttu-id="09874-142">要件</span><span class="sxs-lookup"><span data-stu-id="09874-142">Requirement</span></span>| <span data-ttu-id="09874-143">値</span><span class="sxs-lookup"><span data-stu-id="09874-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="09874-144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="09874-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09874-145">1.0</span><span class="sxs-lookup"><span data-stu-id="09874-145">1.0</span></span>|
|[<span data-ttu-id="09874-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="09874-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09874-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09874-147">ReadItem</span></span>|
|[<span data-ttu-id="09874-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="09874-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="09874-149">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="09874-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09874-150">例</span><span class="sxs-lookup"><span data-stu-id="09874-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="09874-151">timeZone: String</span><span class="sxs-lookup"><span data-stu-id="09874-151">timeZone: String</span></span>

<span data-ttu-id="09874-152">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="09874-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="09874-153">型</span><span class="sxs-lookup"><span data-stu-id="09874-153">Type</span></span>

*   <span data-ttu-id="09874-154">String</span><span class="sxs-lookup"><span data-stu-id="09874-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="09874-155">要件</span><span class="sxs-lookup"><span data-stu-id="09874-155">Requirements</span></span>

|<span data-ttu-id="09874-156">要件</span><span class="sxs-lookup"><span data-stu-id="09874-156">Requirement</span></span>| <span data-ttu-id="09874-157">値</span><span class="sxs-lookup"><span data-stu-id="09874-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="09874-158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="09874-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09874-159">1.0</span><span class="sxs-lookup"><span data-stu-id="09874-159">1.0</span></span>|
|[<span data-ttu-id="09874-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="09874-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09874-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09874-161">ReadItem</span></span>|
|[<span data-ttu-id="09874-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="09874-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="09874-163">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="09874-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09874-164">例</span><span class="sxs-lookup"><span data-stu-id="09874-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
