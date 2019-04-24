---
title: Office.-mailbox-要件セット1.5
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: fc20497cc8df8d091ba0195f7dca9b283ff4d1c2
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451823"
---
# <a name="userprofile"></a><span data-ttu-id="819bb-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="819bb-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="819bb-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="819bb-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="819bb-104">要件</span><span class="sxs-lookup"><span data-stu-id="819bb-104">Requirements</span></span>

|<span data-ttu-id="819bb-105">要件</span><span class="sxs-lookup"><span data-stu-id="819bb-105">Requirement</span></span>| <span data-ttu-id="819bb-106">値</span><span class="sxs-lookup"><span data-stu-id="819bb-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="819bb-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="819bb-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="819bb-108">1.0</span><span class="sxs-lookup"><span data-stu-id="819bb-108">1.0</span></span>|
|[<span data-ttu-id="819bb-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="819bb-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="819bb-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="819bb-110">ReadItem</span></span>|
|[<span data-ttu-id="819bb-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="819bb-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="819bb-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="819bb-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="819bb-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="819bb-113">Members and methods</span></span>

| <span data-ttu-id="819bb-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="819bb-114">Member</span></span> | <span data-ttu-id="819bb-115">種類</span><span class="sxs-lookup"><span data-stu-id="819bb-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="819bb-116">displayName</span><span class="sxs-lookup"><span data-stu-id="819bb-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="819bb-117">Member</span><span class="sxs-lookup"><span data-stu-id="819bb-117">Member</span></span> |
| [<span data-ttu-id="819bb-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="819bb-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="819bb-119">Member</span><span class="sxs-lookup"><span data-stu-id="819bb-119">Member</span></span> |
| [<span data-ttu-id="819bb-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="819bb-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="819bb-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="819bb-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="819bb-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="819bb-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="819bb-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="819bb-123">displayName :String</span></span>

<span data-ttu-id="819bb-124">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="819bb-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="819bb-125">型</span><span class="sxs-lookup"><span data-stu-id="819bb-125">Type</span></span>

*   <span data-ttu-id="819bb-126">String</span><span class="sxs-lookup"><span data-stu-id="819bb-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="819bb-127">要件</span><span class="sxs-lookup"><span data-stu-id="819bb-127">Requirements</span></span>

|<span data-ttu-id="819bb-128">要件</span><span class="sxs-lookup"><span data-stu-id="819bb-128">Requirement</span></span>| <span data-ttu-id="819bb-129">値</span><span class="sxs-lookup"><span data-stu-id="819bb-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="819bb-130">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="819bb-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="819bb-131">1.0</span><span class="sxs-lookup"><span data-stu-id="819bb-131">1.0</span></span>|
|[<span data-ttu-id="819bb-132">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="819bb-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="819bb-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="819bb-133">ReadItem</span></span>|
|[<span data-ttu-id="819bb-134">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="819bb-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="819bb-135">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="819bb-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="819bb-136">例</span><span class="sxs-lookup"><span data-stu-id="819bb-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="819bb-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="819bb-137">emailAddress :String</span></span>

<span data-ttu-id="819bb-138">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="819bb-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="819bb-139">型</span><span class="sxs-lookup"><span data-stu-id="819bb-139">Type</span></span>

*   <span data-ttu-id="819bb-140">String</span><span class="sxs-lookup"><span data-stu-id="819bb-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="819bb-141">要件</span><span class="sxs-lookup"><span data-stu-id="819bb-141">Requirements</span></span>

|<span data-ttu-id="819bb-142">要件</span><span class="sxs-lookup"><span data-stu-id="819bb-142">Requirement</span></span>| <span data-ttu-id="819bb-143">値</span><span class="sxs-lookup"><span data-stu-id="819bb-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="819bb-144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="819bb-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="819bb-145">1.0</span><span class="sxs-lookup"><span data-stu-id="819bb-145">1.0</span></span>|
|[<span data-ttu-id="819bb-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="819bb-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="819bb-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="819bb-147">ReadItem</span></span>|
|[<span data-ttu-id="819bb-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="819bb-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="819bb-149">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="819bb-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="819bb-150">例</span><span class="sxs-lookup"><span data-stu-id="819bb-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="819bb-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="819bb-151">timeZone :String</span></span>

<span data-ttu-id="819bb-152">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="819bb-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="819bb-153">型</span><span class="sxs-lookup"><span data-stu-id="819bb-153">Type</span></span>

*   <span data-ttu-id="819bb-154">String</span><span class="sxs-lookup"><span data-stu-id="819bb-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="819bb-155">要件</span><span class="sxs-lookup"><span data-stu-id="819bb-155">Requirements</span></span>

|<span data-ttu-id="819bb-156">要件</span><span class="sxs-lookup"><span data-stu-id="819bb-156">Requirement</span></span>| <span data-ttu-id="819bb-157">値</span><span class="sxs-lookup"><span data-stu-id="819bb-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="819bb-158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="819bb-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="819bb-159">1.0</span><span class="sxs-lookup"><span data-stu-id="819bb-159">1.0</span></span>|
|[<span data-ttu-id="819bb-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="819bb-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="819bb-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="819bb-161">ReadItem</span></span>|
|[<span data-ttu-id="819bb-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="819bb-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="819bb-163">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="819bb-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="819bb-164">例</span><span class="sxs-lookup"><span data-stu-id="819bb-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
