---
title: Office.-mailbox-要件セット1.4
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 2f8b5bf4b98e55fcc2aa2b58a9a4a7bccc8da51b
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696191"
---
# <a name="userprofile"></a><span data-ttu-id="be730-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="be730-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="be730-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="be730-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="be730-104">要件</span><span class="sxs-lookup"><span data-stu-id="be730-104">Requirements</span></span>

|<span data-ttu-id="be730-105">要件</span><span class="sxs-lookup"><span data-stu-id="be730-105">Requirement</span></span>| <span data-ttu-id="be730-106">値</span><span class="sxs-lookup"><span data-stu-id="be730-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="be730-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="be730-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be730-108">1.0</span><span class="sxs-lookup"><span data-stu-id="be730-108">1.0</span></span>|
|[<span data-ttu-id="be730-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="be730-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be730-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be730-110">ReadItem</span></span>|
|[<span data-ttu-id="be730-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="be730-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be730-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="be730-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="be730-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="be730-113">Members and methods</span></span>

| <span data-ttu-id="be730-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="be730-114">Member</span></span> | <span data-ttu-id="be730-115">種類</span><span class="sxs-lookup"><span data-stu-id="be730-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="be730-116">displayName</span><span class="sxs-lookup"><span data-stu-id="be730-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="be730-117">Member</span><span class="sxs-lookup"><span data-stu-id="be730-117">Member</span></span> |
| [<span data-ttu-id="be730-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="be730-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="be730-119">Member</span><span class="sxs-lookup"><span data-stu-id="be730-119">Member</span></span> |
| [<span data-ttu-id="be730-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="be730-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="be730-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="be730-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="be730-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="be730-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="be730-123">displayName: String</span><span class="sxs-lookup"><span data-stu-id="be730-123">displayName: String</span></span>

<span data-ttu-id="be730-124">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="be730-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="be730-125">型</span><span class="sxs-lookup"><span data-stu-id="be730-125">Type</span></span>

*   <span data-ttu-id="be730-126">String</span><span class="sxs-lookup"><span data-stu-id="be730-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="be730-127">要件</span><span class="sxs-lookup"><span data-stu-id="be730-127">Requirements</span></span>

|<span data-ttu-id="be730-128">要件</span><span class="sxs-lookup"><span data-stu-id="be730-128">Requirement</span></span>| <span data-ttu-id="be730-129">値</span><span class="sxs-lookup"><span data-stu-id="be730-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="be730-130">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="be730-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be730-131">1.0</span><span class="sxs-lookup"><span data-stu-id="be730-131">1.0</span></span>|
|[<span data-ttu-id="be730-132">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="be730-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be730-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be730-133">ReadItem</span></span>|
|[<span data-ttu-id="be730-134">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="be730-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be730-135">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="be730-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be730-136">例</span><span class="sxs-lookup"><span data-stu-id="be730-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="be730-137">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="be730-137">emailAddress: String</span></span>

<span data-ttu-id="be730-138">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="be730-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="be730-139">型</span><span class="sxs-lookup"><span data-stu-id="be730-139">Type</span></span>

*   <span data-ttu-id="be730-140">String</span><span class="sxs-lookup"><span data-stu-id="be730-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="be730-141">要件</span><span class="sxs-lookup"><span data-stu-id="be730-141">Requirements</span></span>

|<span data-ttu-id="be730-142">要件</span><span class="sxs-lookup"><span data-stu-id="be730-142">Requirement</span></span>| <span data-ttu-id="be730-143">値</span><span class="sxs-lookup"><span data-stu-id="be730-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="be730-144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="be730-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be730-145">1.0</span><span class="sxs-lookup"><span data-stu-id="be730-145">1.0</span></span>|
|[<span data-ttu-id="be730-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="be730-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be730-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be730-147">ReadItem</span></span>|
|[<span data-ttu-id="be730-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="be730-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be730-149">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="be730-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be730-150">例</span><span class="sxs-lookup"><span data-stu-id="be730-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="be730-151">timeZone: String</span><span class="sxs-lookup"><span data-stu-id="be730-151">timeZone: String</span></span>

<span data-ttu-id="be730-152">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="be730-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="be730-153">型</span><span class="sxs-lookup"><span data-stu-id="be730-153">Type</span></span>

*   <span data-ttu-id="be730-154">String</span><span class="sxs-lookup"><span data-stu-id="be730-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="be730-155">要件</span><span class="sxs-lookup"><span data-stu-id="be730-155">Requirements</span></span>

|<span data-ttu-id="be730-156">要件</span><span class="sxs-lookup"><span data-stu-id="be730-156">Requirement</span></span>| <span data-ttu-id="be730-157">値</span><span class="sxs-lookup"><span data-stu-id="be730-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="be730-158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="be730-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be730-159">1.0</span><span class="sxs-lookup"><span data-stu-id="be730-159">1.0</span></span>|
|[<span data-ttu-id="be730-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="be730-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be730-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be730-161">ReadItem</span></span>|
|[<span data-ttu-id="be730-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="be730-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be730-163">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="be730-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be730-164">例</span><span class="sxs-lookup"><span data-stu-id="be730-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
