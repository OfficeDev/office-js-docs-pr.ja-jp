---
title: Office.context.mailbox.userProfile - 要件セット 1.2
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: e5548fa514cff9b452c2747324f11e5df8a06def
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432243"
---
# <a name="userprofile"></a><span data-ttu-id="a4a6e-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="a4a6e-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="a4a6e-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="a4a6e-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="a4a6e-104">要件</span><span class="sxs-lookup"><span data-stu-id="a4a6e-104">Requirements</span></span>

|<span data-ttu-id="a4a6e-105">要件</span><span class="sxs-lookup"><span data-stu-id="a4a6e-105">Requirement</span></span>| <span data-ttu-id="a4a6e-106">値</span><span class="sxs-lookup"><span data-stu-id="a4a6e-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4a6e-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a4a6e-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a4a6e-108">1.0</span><span class="sxs-lookup"><span data-stu-id="a4a6e-108">1.0</span></span>|
|[<span data-ttu-id="a4a6e-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a4a6e-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a4a6e-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a4a6e-110">ReadItem</span></span>|
|[<span data-ttu-id="a4a6e-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a4a6e-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a4a6e-112">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a4a6e-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="a4a6e-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="a4a6e-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="a4a6e-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="a4a6e-114">displayName :String</span></span>

<span data-ttu-id="a4a6e-115">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="a4a6e-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="a4a6e-116">型:</span><span class="sxs-lookup"><span data-stu-id="a4a6e-116">Type:</span></span>

*   <span data-ttu-id="a4a6e-117">String</span><span class="sxs-lookup"><span data-stu-id="a4a6e-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a4a6e-118">要件</span><span class="sxs-lookup"><span data-stu-id="a4a6e-118">Requirements</span></span>

|<span data-ttu-id="a4a6e-119">要件</span><span class="sxs-lookup"><span data-stu-id="a4a6e-119">Requirement</span></span>| <span data-ttu-id="a4a6e-120">値</span><span class="sxs-lookup"><span data-stu-id="a4a6e-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4a6e-121">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a4a6e-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a4a6e-122">1.0</span><span class="sxs-lookup"><span data-stu-id="a4a6e-122">1.0</span></span>|
|[<span data-ttu-id="a4a6e-123">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a4a6e-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a4a6e-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a4a6e-124">ReadItem</span></span>|
|[<span data-ttu-id="a4a6e-125">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a4a6e-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a4a6e-126">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a4a6e-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a4a6e-127">例</span><span class="sxs-lookup"><span data-stu-id="a4a6e-127">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="a4a6e-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="a4a6e-128">emailAddress :String</span></span>

<span data-ttu-id="a4a6e-129">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="a4a6e-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="a4a6e-130">型:</span><span class="sxs-lookup"><span data-stu-id="a4a6e-130">Type:</span></span>

*   <span data-ttu-id="a4a6e-131">String</span><span class="sxs-lookup"><span data-stu-id="a4a6e-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a4a6e-132">要件</span><span class="sxs-lookup"><span data-stu-id="a4a6e-132">Requirements</span></span>

|<span data-ttu-id="a4a6e-133">要件</span><span class="sxs-lookup"><span data-stu-id="a4a6e-133">Requirement</span></span>| <span data-ttu-id="a4a6e-134">値</span><span class="sxs-lookup"><span data-stu-id="a4a6e-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4a6e-135">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a4a6e-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a4a6e-136">1.0</span><span class="sxs-lookup"><span data-stu-id="a4a6e-136">1.0</span></span>|
|[<span data-ttu-id="a4a6e-137">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a4a6e-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a4a6e-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a4a6e-138">ReadItem</span></span>|
|[<span data-ttu-id="a4a6e-139">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a4a6e-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a4a6e-140">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a4a6e-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a4a6e-141">例</span><span class="sxs-lookup"><span data-stu-id="a4a6e-141">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="a4a6e-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="a4a6e-142">timeZone :String</span></span>

<span data-ttu-id="a4a6e-143">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="a4a6e-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="a4a6e-144">型:</span><span class="sxs-lookup"><span data-stu-id="a4a6e-144">Type:</span></span>

*   <span data-ttu-id="a4a6e-145">String</span><span class="sxs-lookup"><span data-stu-id="a4a6e-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a4a6e-146">要件</span><span class="sxs-lookup"><span data-stu-id="a4a6e-146">Requirements</span></span>

|<span data-ttu-id="a4a6e-147">要件</span><span class="sxs-lookup"><span data-stu-id="a4a6e-147">Requirement</span></span>| <span data-ttu-id="a4a6e-148">値</span><span class="sxs-lookup"><span data-stu-id="a4a6e-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4a6e-149">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a4a6e-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a4a6e-150">1.0</span><span class="sxs-lookup"><span data-stu-id="a4a6e-150">1.0</span></span>|
|[<span data-ttu-id="a4a6e-151">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a4a6e-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a4a6e-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a4a6e-152">ReadItem</span></span>|
|[<span data-ttu-id="a4a6e-153">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a4a6e-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a4a6e-154">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a4a6e-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a4a6e-155">例</span><span class="sxs-lookup"><span data-stu-id="a4a6e-155">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```