---
title: Office.context.mailbox.userProfile - 要件セット 1.1
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 312cba4d5aace980b7c9b205899fac51d3da3de5
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433174"
---
# <a name="userprofile"></a><span data-ttu-id="21c8b-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="21c8b-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="21c8b-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="21c8b-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="21c8b-104">要件</span><span class="sxs-lookup"><span data-stu-id="21c8b-104">Requirements</span></span>

|<span data-ttu-id="21c8b-105">要件</span><span class="sxs-lookup"><span data-stu-id="21c8b-105">Requirement</span></span>| <span data-ttu-id="21c8b-106">値</span><span class="sxs-lookup"><span data-stu-id="21c8b-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="21c8b-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="21c8b-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="21c8b-108">1.0</span><span class="sxs-lookup"><span data-stu-id="21c8b-108">1.0</span></span>|
|[<span data-ttu-id="21c8b-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="21c8b-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="21c8b-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21c8b-110">ReadItem</span></span>|
|[<span data-ttu-id="21c8b-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="21c8b-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="21c8b-112">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="21c8b-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="21c8b-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="21c8b-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="21c8b-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="21c8b-114">displayName :String</span></span>

<span data-ttu-id="21c8b-115">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="21c8b-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="21c8b-116">型:</span><span class="sxs-lookup"><span data-stu-id="21c8b-116">Type:</span></span>

*   <span data-ttu-id="21c8b-117">String</span><span class="sxs-lookup"><span data-stu-id="21c8b-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="21c8b-118">要件</span><span class="sxs-lookup"><span data-stu-id="21c8b-118">Requirements</span></span>

|<span data-ttu-id="21c8b-119">要件</span><span class="sxs-lookup"><span data-stu-id="21c8b-119">Requirement</span></span>| <span data-ttu-id="21c8b-120">値</span><span class="sxs-lookup"><span data-stu-id="21c8b-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="21c8b-121">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="21c8b-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="21c8b-122">1.0</span><span class="sxs-lookup"><span data-stu-id="21c8b-122">1.0</span></span>|
|[<span data-ttu-id="21c8b-123">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="21c8b-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="21c8b-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21c8b-124">ReadItem</span></span>|
|[<span data-ttu-id="21c8b-125">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="21c8b-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="21c8b-126">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="21c8b-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="21c8b-127">例</span><span class="sxs-lookup"><span data-stu-id="21c8b-127">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="21c8b-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="21c8b-128">emailAddress :String</span></span>

<span data-ttu-id="21c8b-129">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="21c8b-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="21c8b-130">型:</span><span class="sxs-lookup"><span data-stu-id="21c8b-130">Type:</span></span>

*   <span data-ttu-id="21c8b-131">String</span><span class="sxs-lookup"><span data-stu-id="21c8b-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="21c8b-132">要件</span><span class="sxs-lookup"><span data-stu-id="21c8b-132">Requirements</span></span>

|<span data-ttu-id="21c8b-133">要件</span><span class="sxs-lookup"><span data-stu-id="21c8b-133">Requirement</span></span>| <span data-ttu-id="21c8b-134">値</span><span class="sxs-lookup"><span data-stu-id="21c8b-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="21c8b-135">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="21c8b-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="21c8b-136">1.0</span><span class="sxs-lookup"><span data-stu-id="21c8b-136">1.0</span></span>|
|[<span data-ttu-id="21c8b-137">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="21c8b-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="21c8b-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21c8b-138">ReadItem</span></span>|
|[<span data-ttu-id="21c8b-139">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="21c8b-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="21c8b-140">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="21c8b-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="21c8b-141">例</span><span class="sxs-lookup"><span data-stu-id="21c8b-141">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="21c8b-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="21c8b-142">timeZone :String</span></span>

<span data-ttu-id="21c8b-143">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="21c8b-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="21c8b-144">型:</span><span class="sxs-lookup"><span data-stu-id="21c8b-144">Type:</span></span>

*   <span data-ttu-id="21c8b-145">String</span><span class="sxs-lookup"><span data-stu-id="21c8b-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="21c8b-146">要件</span><span class="sxs-lookup"><span data-stu-id="21c8b-146">Requirements</span></span>

|<span data-ttu-id="21c8b-147">要件</span><span class="sxs-lookup"><span data-stu-id="21c8b-147">Requirement</span></span>| <span data-ttu-id="21c8b-148">値</span><span class="sxs-lookup"><span data-stu-id="21c8b-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="21c8b-149">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="21c8b-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="21c8b-150">1.0</span><span class="sxs-lookup"><span data-stu-id="21c8b-150">1.0</span></span>|
|[<span data-ttu-id="21c8b-151">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="21c8b-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="21c8b-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="21c8b-152">ReadItem</span></span>|
|[<span data-ttu-id="21c8b-153">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="21c8b-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="21c8b-154">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="21c8b-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="21c8b-155">例</span><span class="sxs-lookup"><span data-stu-id="21c8b-155">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```