---
title: Office.context.mailbox.userProfile - 要件セット 1.5
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: e98e88cde184db121e69fdd267dff4e39d887b1f
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067832"
---
# <a name="userprofile"></a><span data-ttu-id="06b87-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="06b87-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="06b87-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="06b87-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="06b87-104">要件</span><span class="sxs-lookup"><span data-stu-id="06b87-104">Requirements</span></span>

|<span data-ttu-id="06b87-105">要件</span><span class="sxs-lookup"><span data-stu-id="06b87-105">Requirement</span></span>| <span data-ttu-id="06b87-106">値</span><span class="sxs-lookup"><span data-stu-id="06b87-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="06b87-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="06b87-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="06b87-108">1.0</span><span class="sxs-lookup"><span data-stu-id="06b87-108">1.0</span></span>|
|[<span data-ttu-id="06b87-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="06b87-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="06b87-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="06b87-110">ReadItem</span></span>|
|[<span data-ttu-id="06b87-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="06b87-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="06b87-112">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="06b87-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="06b87-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="06b87-113">Members and methods</span></span>

| <span data-ttu-id="06b87-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="06b87-114">Member</span></span> | <span data-ttu-id="06b87-115">種類</span><span class="sxs-lookup"><span data-stu-id="06b87-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="06b87-116">displayName</span><span class="sxs-lookup"><span data-stu-id="06b87-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="06b87-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="06b87-117">Member</span></span> |
| [<span data-ttu-id="06b87-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="06b87-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="06b87-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="06b87-119">Member</span></span> |
| [<span data-ttu-id="06b87-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="06b87-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="06b87-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="06b87-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="06b87-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="06b87-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="06b87-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="06b87-123">displayName :String</span></span>

<span data-ttu-id="06b87-124">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="06b87-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="06b87-125">Type</span><span class="sxs-lookup"><span data-stu-id="06b87-125">Type</span></span>

*   <span data-ttu-id="06b87-126">String</span><span class="sxs-lookup"><span data-stu-id="06b87-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="06b87-127">要件</span><span class="sxs-lookup"><span data-stu-id="06b87-127">Requirements</span></span>

|<span data-ttu-id="06b87-128">要件</span><span class="sxs-lookup"><span data-stu-id="06b87-128">Requirement</span></span>| <span data-ttu-id="06b87-129">値</span><span class="sxs-lookup"><span data-stu-id="06b87-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="06b87-130">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="06b87-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="06b87-131">1.0</span><span class="sxs-lookup"><span data-stu-id="06b87-131">1.0</span></span>|
|[<span data-ttu-id="06b87-132">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="06b87-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="06b87-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="06b87-133">ReadItem</span></span>|
|[<span data-ttu-id="06b87-134">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="06b87-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="06b87-135">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="06b87-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="06b87-136">例</span><span class="sxs-lookup"><span data-stu-id="06b87-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="06b87-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="06b87-137">emailAddress :String</span></span>

<span data-ttu-id="06b87-138">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="06b87-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="06b87-139">Type</span><span class="sxs-lookup"><span data-stu-id="06b87-139">Type</span></span>

*   <span data-ttu-id="06b87-140">String</span><span class="sxs-lookup"><span data-stu-id="06b87-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="06b87-141">要件</span><span class="sxs-lookup"><span data-stu-id="06b87-141">Requirements</span></span>

|<span data-ttu-id="06b87-142">要件</span><span class="sxs-lookup"><span data-stu-id="06b87-142">Requirement</span></span>| <span data-ttu-id="06b87-143">値</span><span class="sxs-lookup"><span data-stu-id="06b87-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="06b87-144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="06b87-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="06b87-145">1.0</span><span class="sxs-lookup"><span data-stu-id="06b87-145">1.0</span></span>|
|[<span data-ttu-id="06b87-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="06b87-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="06b87-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="06b87-147">ReadItem</span></span>|
|[<span data-ttu-id="06b87-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="06b87-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="06b87-149">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="06b87-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="06b87-150">例</span><span class="sxs-lookup"><span data-stu-id="06b87-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="06b87-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="06b87-151">timeZone :String</span></span>

<span data-ttu-id="06b87-152">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="06b87-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="06b87-153">Type</span><span class="sxs-lookup"><span data-stu-id="06b87-153">Type</span></span>

*   <span data-ttu-id="06b87-154">String</span><span class="sxs-lookup"><span data-stu-id="06b87-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="06b87-155">要件</span><span class="sxs-lookup"><span data-stu-id="06b87-155">Requirements</span></span>

|<span data-ttu-id="06b87-156">要件</span><span class="sxs-lookup"><span data-stu-id="06b87-156">Requirement</span></span>| <span data-ttu-id="06b87-157">値</span><span class="sxs-lookup"><span data-stu-id="06b87-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="06b87-158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="06b87-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="06b87-159">1.0</span><span class="sxs-lookup"><span data-stu-id="06b87-159">1.0</span></span>|
|[<span data-ttu-id="06b87-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="06b87-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="06b87-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="06b87-161">ReadItem</span></span>|
|[<span data-ttu-id="06b87-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="06b87-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="06b87-163">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="06b87-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="06b87-164">例</span><span class="sxs-lookup"><span data-stu-id="06b87-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
