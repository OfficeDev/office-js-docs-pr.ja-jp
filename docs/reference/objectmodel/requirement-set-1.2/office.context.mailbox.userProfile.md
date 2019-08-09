---
title: Office.-mailbox-要件セット1.2
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 7258195e7ec0ef2432723d0f32f3d9ef1a3acf2b
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268685"
---
# <a name="userprofile"></a><span data-ttu-id="f6384-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="f6384-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="f6384-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="f6384-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6384-104">要件</span><span class="sxs-lookup"><span data-stu-id="f6384-104">Requirements</span></span>

|<span data-ttu-id="f6384-105">要件</span><span class="sxs-lookup"><span data-stu-id="f6384-105">Requirement</span></span>| <span data-ttu-id="f6384-106">値</span><span class="sxs-lookup"><span data-stu-id="f6384-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6384-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f6384-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f6384-108">1.0</span><span class="sxs-lookup"><span data-stu-id="f6384-108">1.0</span></span>|
|[<span data-ttu-id="f6384-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f6384-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f6384-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6384-110">ReadItem</span></span>|
|[<span data-ttu-id="f6384-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f6384-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f6384-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f6384-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f6384-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="f6384-113">Members and methods</span></span>

| <span data-ttu-id="f6384-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="f6384-114">Member</span></span> | <span data-ttu-id="f6384-115">種類</span><span class="sxs-lookup"><span data-stu-id="f6384-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f6384-116">displayName</span><span class="sxs-lookup"><span data-stu-id="f6384-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="f6384-117">Member</span><span class="sxs-lookup"><span data-stu-id="f6384-117">Member</span></span> |
| [<span data-ttu-id="f6384-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="f6384-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="f6384-119">Member</span><span class="sxs-lookup"><span data-stu-id="f6384-119">Member</span></span> |
| [<span data-ttu-id="f6384-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="f6384-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="f6384-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="f6384-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="f6384-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="f6384-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="f6384-123">displayName: String</span><span class="sxs-lookup"><span data-stu-id="f6384-123">displayName: String</span></span>

<span data-ttu-id="f6384-124">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="f6384-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="f6384-125">型</span><span class="sxs-lookup"><span data-stu-id="f6384-125">Type</span></span>

*   <span data-ttu-id="f6384-126">String</span><span class="sxs-lookup"><span data-stu-id="f6384-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6384-127">要件</span><span class="sxs-lookup"><span data-stu-id="f6384-127">Requirements</span></span>

|<span data-ttu-id="f6384-128">要件</span><span class="sxs-lookup"><span data-stu-id="f6384-128">Requirement</span></span>| <span data-ttu-id="f6384-129">値</span><span class="sxs-lookup"><span data-stu-id="f6384-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6384-130">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f6384-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f6384-131">1.0</span><span class="sxs-lookup"><span data-stu-id="f6384-131">1.0</span></span>|
|[<span data-ttu-id="f6384-132">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f6384-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f6384-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6384-133">ReadItem</span></span>|
|[<span data-ttu-id="f6384-134">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f6384-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f6384-135">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f6384-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6384-136">例</span><span class="sxs-lookup"><span data-stu-id="f6384-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

#### <a name="emailaddress-string"></a><span data-ttu-id="f6384-137">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="f6384-137">emailAddress: String</span></span>

<span data-ttu-id="f6384-138">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="f6384-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="f6384-139">型</span><span class="sxs-lookup"><span data-stu-id="f6384-139">Type</span></span>

*   <span data-ttu-id="f6384-140">String</span><span class="sxs-lookup"><span data-stu-id="f6384-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6384-141">要件</span><span class="sxs-lookup"><span data-stu-id="f6384-141">Requirements</span></span>

|<span data-ttu-id="f6384-142">要件</span><span class="sxs-lookup"><span data-stu-id="f6384-142">Requirement</span></span>| <span data-ttu-id="f6384-143">値</span><span class="sxs-lookup"><span data-stu-id="f6384-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6384-144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f6384-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f6384-145">1.0</span><span class="sxs-lookup"><span data-stu-id="f6384-145">1.0</span></span>|
|[<span data-ttu-id="f6384-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f6384-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f6384-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6384-147">ReadItem</span></span>|
|[<span data-ttu-id="f6384-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f6384-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f6384-149">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f6384-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6384-150">例</span><span class="sxs-lookup"><span data-stu-id="f6384-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

#### <a name="timezone-string"></a><span data-ttu-id="f6384-151">timeZone: String</span><span class="sxs-lookup"><span data-stu-id="f6384-151">timeZone: String</span></span>

<span data-ttu-id="f6384-152">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="f6384-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="f6384-153">型</span><span class="sxs-lookup"><span data-stu-id="f6384-153">Type</span></span>

*   <span data-ttu-id="f6384-154">String</span><span class="sxs-lookup"><span data-stu-id="f6384-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f6384-155">要件</span><span class="sxs-lookup"><span data-stu-id="f6384-155">Requirements</span></span>

|<span data-ttu-id="f6384-156">要件</span><span class="sxs-lookup"><span data-stu-id="f6384-156">Requirement</span></span>| <span data-ttu-id="f6384-157">値</span><span class="sxs-lookup"><span data-stu-id="f6384-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="f6384-158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f6384-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f6384-159">1.0</span><span class="sxs-lookup"><span data-stu-id="f6384-159">1.0</span></span>|
|[<span data-ttu-id="f6384-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f6384-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f6384-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f6384-161">ReadItem</span></span>|
|[<span data-ttu-id="f6384-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f6384-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f6384-163">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f6384-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f6384-164">例</span><span class="sxs-lookup"><span data-stu-id="f6384-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
