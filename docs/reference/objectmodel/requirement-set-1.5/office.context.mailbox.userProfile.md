---
title: Office.-mailbox-要件セット1.5
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 993fad674fcc616483ac927619e7ca64d81b7326
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696093"
---
# <a name="userprofile"></a><span data-ttu-id="f435c-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="f435c-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="f435c-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="f435c-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="f435c-104">要件</span><span class="sxs-lookup"><span data-stu-id="f435c-104">Requirements</span></span>

|<span data-ttu-id="f435c-105">要件</span><span class="sxs-lookup"><span data-stu-id="f435c-105">Requirement</span></span>| <span data-ttu-id="f435c-106">値</span><span class="sxs-lookup"><span data-stu-id="f435c-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="f435c-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f435c-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f435c-108">1.0</span><span class="sxs-lookup"><span data-stu-id="f435c-108">1.0</span></span>|
|[<span data-ttu-id="f435c-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f435c-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f435c-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f435c-110">ReadItem</span></span>|
|[<span data-ttu-id="f435c-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f435c-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f435c-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f435c-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f435c-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="f435c-113">Members and methods</span></span>

| <span data-ttu-id="f435c-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="f435c-114">Member</span></span> | <span data-ttu-id="f435c-115">種類</span><span class="sxs-lookup"><span data-stu-id="f435c-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f435c-116">displayName</span><span class="sxs-lookup"><span data-stu-id="f435c-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="f435c-117">Member</span><span class="sxs-lookup"><span data-stu-id="f435c-117">Member</span></span> |
| [<span data-ttu-id="f435c-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="f435c-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="f435c-119">Member</span><span class="sxs-lookup"><span data-stu-id="f435c-119">Member</span></span> |
| [<span data-ttu-id="f435c-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="f435c-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="f435c-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="f435c-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="f435c-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="f435c-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="f435c-123">displayName: String</span><span class="sxs-lookup"><span data-stu-id="f435c-123">displayName: String</span></span>

<span data-ttu-id="f435c-124">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="f435c-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="f435c-125">型</span><span class="sxs-lookup"><span data-stu-id="f435c-125">Type</span></span>

*   <span data-ttu-id="f435c-126">String</span><span class="sxs-lookup"><span data-stu-id="f435c-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f435c-127">要件</span><span class="sxs-lookup"><span data-stu-id="f435c-127">Requirements</span></span>

|<span data-ttu-id="f435c-128">要件</span><span class="sxs-lookup"><span data-stu-id="f435c-128">Requirement</span></span>| <span data-ttu-id="f435c-129">値</span><span class="sxs-lookup"><span data-stu-id="f435c-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="f435c-130">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f435c-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f435c-131">1.0</span><span class="sxs-lookup"><span data-stu-id="f435c-131">1.0</span></span>|
|[<span data-ttu-id="f435c-132">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f435c-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f435c-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f435c-133">ReadItem</span></span>|
|[<span data-ttu-id="f435c-134">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f435c-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f435c-135">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f435c-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f435c-136">例</span><span class="sxs-lookup"><span data-stu-id="f435c-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="f435c-137">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="f435c-137">emailAddress: String</span></span>

<span data-ttu-id="f435c-138">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="f435c-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="f435c-139">型</span><span class="sxs-lookup"><span data-stu-id="f435c-139">Type</span></span>

*   <span data-ttu-id="f435c-140">String</span><span class="sxs-lookup"><span data-stu-id="f435c-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f435c-141">要件</span><span class="sxs-lookup"><span data-stu-id="f435c-141">Requirements</span></span>

|<span data-ttu-id="f435c-142">要件</span><span class="sxs-lookup"><span data-stu-id="f435c-142">Requirement</span></span>| <span data-ttu-id="f435c-143">値</span><span class="sxs-lookup"><span data-stu-id="f435c-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="f435c-144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f435c-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f435c-145">1.0</span><span class="sxs-lookup"><span data-stu-id="f435c-145">1.0</span></span>|
|[<span data-ttu-id="f435c-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f435c-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f435c-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f435c-147">ReadItem</span></span>|
|[<span data-ttu-id="f435c-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f435c-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f435c-149">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f435c-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f435c-150">例</span><span class="sxs-lookup"><span data-stu-id="f435c-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="f435c-151">timeZone: String</span><span class="sxs-lookup"><span data-stu-id="f435c-151">timeZone: String</span></span>

<span data-ttu-id="f435c-152">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="f435c-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="f435c-153">型</span><span class="sxs-lookup"><span data-stu-id="f435c-153">Type</span></span>

*   <span data-ttu-id="f435c-154">String</span><span class="sxs-lookup"><span data-stu-id="f435c-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f435c-155">要件</span><span class="sxs-lookup"><span data-stu-id="f435c-155">Requirements</span></span>

|<span data-ttu-id="f435c-156">要件</span><span class="sxs-lookup"><span data-stu-id="f435c-156">Requirement</span></span>| <span data-ttu-id="f435c-157">値</span><span class="sxs-lookup"><span data-stu-id="f435c-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="f435c-158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f435c-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f435c-159">1.0</span><span class="sxs-lookup"><span data-stu-id="f435c-159">1.0</span></span>|
|[<span data-ttu-id="f435c-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f435c-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f435c-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f435c-161">ReadItem</span></span>|
|[<span data-ttu-id="f435c-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f435c-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f435c-163">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f435c-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f435c-164">例</span><span class="sxs-lookup"><span data-stu-id="f435c-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
