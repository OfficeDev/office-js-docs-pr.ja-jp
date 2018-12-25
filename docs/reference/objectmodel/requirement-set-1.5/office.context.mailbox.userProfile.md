---
title: Office.context.mailbox.userProfile - 要件セット 1.5
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 748daf4d14aae1d14560d29e1d76eeea09830573
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432719"
---
# <a name="userprofile"></a><span data-ttu-id="3db85-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="3db85-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="3db85-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="3db85-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="3db85-104">要件</span><span class="sxs-lookup"><span data-stu-id="3db85-104">Requirements</span></span>

|<span data-ttu-id="3db85-105">要件</span><span class="sxs-lookup"><span data-stu-id="3db85-105">Requirement</span></span>| <span data-ttu-id="3db85-106">値</span><span class="sxs-lookup"><span data-stu-id="3db85-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="3db85-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3db85-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3db85-108">1.0</span><span class="sxs-lookup"><span data-stu-id="3db85-108">1.0</span></span>|
|[<span data-ttu-id="3db85-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="3db85-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3db85-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3db85-110">ReadItem</span></span>|
|[<span data-ttu-id="3db85-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3db85-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3db85-112">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3db85-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="3db85-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="3db85-113">Members and methods</span></span>

| <span data-ttu-id="3db85-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="3db85-114">Member</span></span> | <span data-ttu-id="3db85-115">種類</span><span class="sxs-lookup"><span data-stu-id="3db85-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="3db85-116">displayName</span><span class="sxs-lookup"><span data-stu-id="3db85-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="3db85-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="3db85-117">Member</span></span> |
| [<span data-ttu-id="3db85-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="3db85-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="3db85-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="3db85-119">Member</span></span> |
| [<span data-ttu-id="3db85-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="3db85-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="3db85-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="3db85-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="3db85-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="3db85-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="3db85-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="3db85-123">displayName :String</span></span>

<span data-ttu-id="3db85-124">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="3db85-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="3db85-125">型:</span><span class="sxs-lookup"><span data-stu-id="3db85-125">Type:</span></span>

*   <span data-ttu-id="3db85-126">String</span><span class="sxs-lookup"><span data-stu-id="3db85-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3db85-127">要件</span><span class="sxs-lookup"><span data-stu-id="3db85-127">Requirements</span></span>

|<span data-ttu-id="3db85-128">要件</span><span class="sxs-lookup"><span data-stu-id="3db85-128">Requirement</span></span>| <span data-ttu-id="3db85-129">値</span><span class="sxs-lookup"><span data-stu-id="3db85-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="3db85-130">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3db85-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3db85-131">1.0</span><span class="sxs-lookup"><span data-stu-id="3db85-131">1.0</span></span>|
|[<span data-ttu-id="3db85-132">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="3db85-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3db85-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3db85-133">ReadItem</span></span>|
|[<span data-ttu-id="3db85-134">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3db85-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3db85-135">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3db85-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3db85-136">例</span><span class="sxs-lookup"><span data-stu-id="3db85-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="3db85-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="3db85-137">emailAddress :String</span></span>

<span data-ttu-id="3db85-138">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="3db85-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="3db85-139">型:</span><span class="sxs-lookup"><span data-stu-id="3db85-139">Type:</span></span>

*   <span data-ttu-id="3db85-140">String</span><span class="sxs-lookup"><span data-stu-id="3db85-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3db85-141">要件</span><span class="sxs-lookup"><span data-stu-id="3db85-141">Requirements</span></span>

|<span data-ttu-id="3db85-142">要件</span><span class="sxs-lookup"><span data-stu-id="3db85-142">Requirement</span></span>| <span data-ttu-id="3db85-143">値</span><span class="sxs-lookup"><span data-stu-id="3db85-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="3db85-144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3db85-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3db85-145">1.0</span><span class="sxs-lookup"><span data-stu-id="3db85-145">1.0</span></span>|
|[<span data-ttu-id="3db85-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="3db85-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3db85-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3db85-147">ReadItem</span></span>|
|[<span data-ttu-id="3db85-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3db85-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3db85-149">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3db85-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3db85-150">例</span><span class="sxs-lookup"><span data-stu-id="3db85-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="3db85-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="3db85-151">timeZone :String</span></span>

<span data-ttu-id="3db85-152">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="3db85-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="3db85-153">型:</span><span class="sxs-lookup"><span data-stu-id="3db85-153">Type:</span></span>

*   <span data-ttu-id="3db85-154">String</span><span class="sxs-lookup"><span data-stu-id="3db85-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3db85-155">要件</span><span class="sxs-lookup"><span data-stu-id="3db85-155">Requirements</span></span>

|<span data-ttu-id="3db85-156">要件</span><span class="sxs-lookup"><span data-stu-id="3db85-156">Requirement</span></span>| <span data-ttu-id="3db85-157">値</span><span class="sxs-lookup"><span data-stu-id="3db85-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="3db85-158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3db85-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3db85-159">1.0</span><span class="sxs-lookup"><span data-stu-id="3db85-159">1.0</span></span>|
|[<span data-ttu-id="3db85-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="3db85-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3db85-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3db85-161">ReadItem</span></span>|
|[<span data-ttu-id="3db85-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3db85-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3db85-163">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3db85-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3db85-164">例</span><span class="sxs-lookup"><span data-stu-id="3db85-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```