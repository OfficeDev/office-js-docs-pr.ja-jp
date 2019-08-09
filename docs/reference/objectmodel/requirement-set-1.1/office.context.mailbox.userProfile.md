---
title: Office.-mailbox-要件セット1.1
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: af9a7f790f56124a86af08567690452b7f497408
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268489"
---
# <a name="userprofile"></a><span data-ttu-id="ab85d-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="ab85d-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="ab85d-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="ab85d-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab85d-104">要件</span><span class="sxs-lookup"><span data-stu-id="ab85d-104">Requirements</span></span>

|<span data-ttu-id="ab85d-105">要件</span><span class="sxs-lookup"><span data-stu-id="ab85d-105">Requirement</span></span>| <span data-ttu-id="ab85d-106">値</span><span class="sxs-lookup"><span data-stu-id="ab85d-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab85d-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ab85d-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab85d-108">1.0</span><span class="sxs-lookup"><span data-stu-id="ab85d-108">1.0</span></span>|
|[<span data-ttu-id="ab85d-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ab85d-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab85d-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab85d-110">ReadItem</span></span>|
|[<span data-ttu-id="ab85d-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ab85d-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ab85d-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ab85d-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ab85d-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="ab85d-113">Members and methods</span></span>

| <span data-ttu-id="ab85d-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="ab85d-114">Member</span></span> | <span data-ttu-id="ab85d-115">種類</span><span class="sxs-lookup"><span data-stu-id="ab85d-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ab85d-116">displayName</span><span class="sxs-lookup"><span data-stu-id="ab85d-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="ab85d-117">Member</span><span class="sxs-lookup"><span data-stu-id="ab85d-117">Member</span></span> |
| [<span data-ttu-id="ab85d-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="ab85d-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="ab85d-119">Member</span><span class="sxs-lookup"><span data-stu-id="ab85d-119">Member</span></span> |
| [<span data-ttu-id="ab85d-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="ab85d-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="ab85d-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="ab85d-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="ab85d-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="ab85d-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="ab85d-123">displayName: String</span><span class="sxs-lookup"><span data-stu-id="ab85d-123">displayName: String</span></span>

<span data-ttu-id="ab85d-124">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="ab85d-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="ab85d-125">型</span><span class="sxs-lookup"><span data-stu-id="ab85d-125">Type</span></span>

*   <span data-ttu-id="ab85d-126">String</span><span class="sxs-lookup"><span data-stu-id="ab85d-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab85d-127">要件</span><span class="sxs-lookup"><span data-stu-id="ab85d-127">Requirements</span></span>

|<span data-ttu-id="ab85d-128">要件</span><span class="sxs-lookup"><span data-stu-id="ab85d-128">Requirement</span></span>| <span data-ttu-id="ab85d-129">値</span><span class="sxs-lookup"><span data-stu-id="ab85d-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab85d-130">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ab85d-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab85d-131">1.0</span><span class="sxs-lookup"><span data-stu-id="ab85d-131">1.0</span></span>|
|[<span data-ttu-id="ab85d-132">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ab85d-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab85d-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab85d-133">ReadItem</span></span>|
|[<span data-ttu-id="ab85d-134">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ab85d-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ab85d-135">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ab85d-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab85d-136">例</span><span class="sxs-lookup"><span data-stu-id="ab85d-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

#### <a name="emailaddress-string"></a><span data-ttu-id="ab85d-137">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="ab85d-137">emailAddress: String</span></span>

<span data-ttu-id="ab85d-138">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="ab85d-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="ab85d-139">型</span><span class="sxs-lookup"><span data-stu-id="ab85d-139">Type</span></span>

*   <span data-ttu-id="ab85d-140">String</span><span class="sxs-lookup"><span data-stu-id="ab85d-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab85d-141">要件</span><span class="sxs-lookup"><span data-stu-id="ab85d-141">Requirements</span></span>

|<span data-ttu-id="ab85d-142">要件</span><span class="sxs-lookup"><span data-stu-id="ab85d-142">Requirement</span></span>| <span data-ttu-id="ab85d-143">値</span><span class="sxs-lookup"><span data-stu-id="ab85d-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab85d-144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ab85d-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab85d-145">1.0</span><span class="sxs-lookup"><span data-stu-id="ab85d-145">1.0</span></span>|
|[<span data-ttu-id="ab85d-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ab85d-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab85d-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab85d-147">ReadItem</span></span>|
|[<span data-ttu-id="ab85d-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ab85d-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ab85d-149">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ab85d-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab85d-150">例</span><span class="sxs-lookup"><span data-stu-id="ab85d-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

#### <a name="timezone-string"></a><span data-ttu-id="ab85d-151">timeZone: String</span><span class="sxs-lookup"><span data-stu-id="ab85d-151">timeZone: String</span></span>

<span data-ttu-id="ab85d-152">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="ab85d-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="ab85d-153">型</span><span class="sxs-lookup"><span data-stu-id="ab85d-153">Type</span></span>

*   <span data-ttu-id="ab85d-154">String</span><span class="sxs-lookup"><span data-stu-id="ab85d-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab85d-155">要件</span><span class="sxs-lookup"><span data-stu-id="ab85d-155">Requirements</span></span>

|<span data-ttu-id="ab85d-156">要件</span><span class="sxs-lookup"><span data-stu-id="ab85d-156">Requirement</span></span>| <span data-ttu-id="ab85d-157">値</span><span class="sxs-lookup"><span data-stu-id="ab85d-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab85d-158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ab85d-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab85d-159">1.0</span><span class="sxs-lookup"><span data-stu-id="ab85d-159">1.0</span></span>|
|[<span data-ttu-id="ab85d-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ab85d-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab85d-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab85d-161">ReadItem</span></span>|
|[<span data-ttu-id="ab85d-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ab85d-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ab85d-163">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ab85d-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab85d-164">例</span><span class="sxs-lookup"><span data-stu-id="ab85d-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
