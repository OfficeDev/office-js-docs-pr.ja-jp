---
title: Office.context.mailbox.userProfile - 要件セット 1.4
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 7facc0ea555dca7d6784a09f798c3d8fa25f2731
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067847"
---
# <a name="userprofile"></a><span data-ttu-id="65703-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="65703-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="65703-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="65703-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="65703-104">要件</span><span class="sxs-lookup"><span data-stu-id="65703-104">Requirements</span></span>

|<span data-ttu-id="65703-105">要件</span><span class="sxs-lookup"><span data-stu-id="65703-105">Requirement</span></span>| <span data-ttu-id="65703-106">値</span><span class="sxs-lookup"><span data-stu-id="65703-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="65703-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="65703-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="65703-108">1.0</span><span class="sxs-lookup"><span data-stu-id="65703-108">1.0</span></span>|
|[<span data-ttu-id="65703-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="65703-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="65703-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="65703-110">ReadItem</span></span>|
|[<span data-ttu-id="65703-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="65703-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="65703-112">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="65703-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="65703-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="65703-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="65703-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="65703-114">displayName :String</span></span>

<span data-ttu-id="65703-115">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="65703-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="65703-116">Type</span><span class="sxs-lookup"><span data-stu-id="65703-116">Type</span></span>

*   <span data-ttu-id="65703-117">String</span><span class="sxs-lookup"><span data-stu-id="65703-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="65703-118">要件</span><span class="sxs-lookup"><span data-stu-id="65703-118">Requirements</span></span>

|<span data-ttu-id="65703-119">要件</span><span class="sxs-lookup"><span data-stu-id="65703-119">Requirement</span></span>| <span data-ttu-id="65703-120">値</span><span class="sxs-lookup"><span data-stu-id="65703-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="65703-121">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="65703-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="65703-122">1.0</span><span class="sxs-lookup"><span data-stu-id="65703-122">1.0</span></span>|
|[<span data-ttu-id="65703-123">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="65703-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="65703-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="65703-124">ReadItem</span></span>|
|[<span data-ttu-id="65703-125">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="65703-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="65703-126">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="65703-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="65703-127">例</span><span class="sxs-lookup"><span data-stu-id="65703-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="65703-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="65703-128">emailAddress :String</span></span>

<span data-ttu-id="65703-129">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="65703-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="65703-130">Type</span><span class="sxs-lookup"><span data-stu-id="65703-130">Type</span></span>

*   <span data-ttu-id="65703-131">String</span><span class="sxs-lookup"><span data-stu-id="65703-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="65703-132">要件</span><span class="sxs-lookup"><span data-stu-id="65703-132">Requirements</span></span>

|<span data-ttu-id="65703-133">要件</span><span class="sxs-lookup"><span data-stu-id="65703-133">Requirement</span></span>| <span data-ttu-id="65703-134">値</span><span class="sxs-lookup"><span data-stu-id="65703-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="65703-135">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="65703-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="65703-136">1.0</span><span class="sxs-lookup"><span data-stu-id="65703-136">1.0</span></span>|
|[<span data-ttu-id="65703-137">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="65703-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="65703-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="65703-138">ReadItem</span></span>|
|[<span data-ttu-id="65703-139">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="65703-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="65703-140">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="65703-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="65703-141">例</span><span class="sxs-lookup"><span data-stu-id="65703-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="65703-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="65703-142">timeZone :String</span></span>

<span data-ttu-id="65703-143">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="65703-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="65703-144">Type</span><span class="sxs-lookup"><span data-stu-id="65703-144">Type</span></span>

*   <span data-ttu-id="65703-145">String</span><span class="sxs-lookup"><span data-stu-id="65703-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="65703-146">要件</span><span class="sxs-lookup"><span data-stu-id="65703-146">Requirements</span></span>

|<span data-ttu-id="65703-147">要件</span><span class="sxs-lookup"><span data-stu-id="65703-147">Requirement</span></span>| <span data-ttu-id="65703-148">値</span><span class="sxs-lookup"><span data-stu-id="65703-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="65703-149">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="65703-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="65703-150">1.0</span><span class="sxs-lookup"><span data-stu-id="65703-150">1.0</span></span>|
|[<span data-ttu-id="65703-151">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="65703-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="65703-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="65703-152">ReadItem</span></span>|
|[<span data-ttu-id="65703-153">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="65703-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="65703-154">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="65703-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="65703-155">例</span><span class="sxs-lookup"><span data-stu-id="65703-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
