---
title: Office.-mailbox-要件セット1.1
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 06492623e0b9ab16792d6b23dfaeb27d99125ff1
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696401"
---
# <a name="userprofile"></a><span data-ttu-id="1df57-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="1df57-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="1df57-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="1df57-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="1df57-104">要件</span><span class="sxs-lookup"><span data-stu-id="1df57-104">Requirements</span></span>

|<span data-ttu-id="1df57-105">要件</span><span class="sxs-lookup"><span data-stu-id="1df57-105">Requirement</span></span>| <span data-ttu-id="1df57-106">値</span><span class="sxs-lookup"><span data-stu-id="1df57-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="1df57-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1df57-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1df57-108">1.0</span><span class="sxs-lookup"><span data-stu-id="1df57-108">1.0</span></span>|
|[<span data-ttu-id="1df57-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1df57-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1df57-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1df57-110">ReadItem</span></span>|
|[<span data-ttu-id="1df57-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1df57-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1df57-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1df57-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1df57-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="1df57-113">Members and methods</span></span>

| <span data-ttu-id="1df57-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="1df57-114">Member</span></span> | <span data-ttu-id="1df57-115">種類</span><span class="sxs-lookup"><span data-stu-id="1df57-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1df57-116">displayName</span><span class="sxs-lookup"><span data-stu-id="1df57-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="1df57-117">Member</span><span class="sxs-lookup"><span data-stu-id="1df57-117">Member</span></span> |
| [<span data-ttu-id="1df57-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="1df57-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="1df57-119">Member</span><span class="sxs-lookup"><span data-stu-id="1df57-119">Member</span></span> |
| [<span data-ttu-id="1df57-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="1df57-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="1df57-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="1df57-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="1df57-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="1df57-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="1df57-123">displayName: String</span><span class="sxs-lookup"><span data-stu-id="1df57-123">displayName: String</span></span>

<span data-ttu-id="1df57-124">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="1df57-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="1df57-125">型</span><span class="sxs-lookup"><span data-stu-id="1df57-125">Type</span></span>

*   <span data-ttu-id="1df57-126">String</span><span class="sxs-lookup"><span data-stu-id="1df57-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1df57-127">要件</span><span class="sxs-lookup"><span data-stu-id="1df57-127">Requirements</span></span>

|<span data-ttu-id="1df57-128">要件</span><span class="sxs-lookup"><span data-stu-id="1df57-128">Requirement</span></span>| <span data-ttu-id="1df57-129">値</span><span class="sxs-lookup"><span data-stu-id="1df57-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="1df57-130">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1df57-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1df57-131">1.0</span><span class="sxs-lookup"><span data-stu-id="1df57-131">1.0</span></span>|
|[<span data-ttu-id="1df57-132">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1df57-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1df57-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1df57-133">ReadItem</span></span>|
|[<span data-ttu-id="1df57-134">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1df57-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1df57-135">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1df57-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1df57-136">例</span><span class="sxs-lookup"><span data-stu-id="1df57-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="1df57-137">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="1df57-137">emailAddress: String</span></span>

<span data-ttu-id="1df57-138">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="1df57-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="1df57-139">型</span><span class="sxs-lookup"><span data-stu-id="1df57-139">Type</span></span>

*   <span data-ttu-id="1df57-140">String</span><span class="sxs-lookup"><span data-stu-id="1df57-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1df57-141">要件</span><span class="sxs-lookup"><span data-stu-id="1df57-141">Requirements</span></span>

|<span data-ttu-id="1df57-142">要件</span><span class="sxs-lookup"><span data-stu-id="1df57-142">Requirement</span></span>| <span data-ttu-id="1df57-143">値</span><span class="sxs-lookup"><span data-stu-id="1df57-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="1df57-144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1df57-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1df57-145">1.0</span><span class="sxs-lookup"><span data-stu-id="1df57-145">1.0</span></span>|
|[<span data-ttu-id="1df57-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1df57-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1df57-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1df57-147">ReadItem</span></span>|
|[<span data-ttu-id="1df57-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1df57-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1df57-149">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1df57-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1df57-150">例</span><span class="sxs-lookup"><span data-stu-id="1df57-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="1df57-151">timeZone: String</span><span class="sxs-lookup"><span data-stu-id="1df57-151">timeZone: String</span></span>

<span data-ttu-id="1df57-152">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="1df57-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="1df57-153">型</span><span class="sxs-lookup"><span data-stu-id="1df57-153">Type</span></span>

*   <span data-ttu-id="1df57-154">String</span><span class="sxs-lookup"><span data-stu-id="1df57-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1df57-155">要件</span><span class="sxs-lookup"><span data-stu-id="1df57-155">Requirements</span></span>

|<span data-ttu-id="1df57-156">要件</span><span class="sxs-lookup"><span data-stu-id="1df57-156">Requirement</span></span>| <span data-ttu-id="1df57-157">値</span><span class="sxs-lookup"><span data-stu-id="1df57-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="1df57-158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1df57-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1df57-159">1.0</span><span class="sxs-lookup"><span data-stu-id="1df57-159">1.0</span></span>|
|[<span data-ttu-id="1df57-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1df57-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1df57-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1df57-161">ReadItem</span></span>|
|[<span data-ttu-id="1df57-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1df57-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1df57-163">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1df57-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1df57-164">例</span><span class="sxs-lookup"><span data-stu-id="1df57-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
