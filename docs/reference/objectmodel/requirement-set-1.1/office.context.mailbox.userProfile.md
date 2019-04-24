---
title: Office.-mailbox-要件セット1.1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 7a10a35887d31a8803d0662eedbe190543d2326a
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451921"
---
# <a name="userprofile"></a><span data-ttu-id="0cb5c-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="0cb5c-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="0cb5c-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="0cb5c-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cb5c-104">要件</span><span class="sxs-lookup"><span data-stu-id="0cb5c-104">Requirements</span></span>

|<span data-ttu-id="0cb5c-105">要件</span><span class="sxs-lookup"><span data-stu-id="0cb5c-105">Requirement</span></span>| <span data-ttu-id="0cb5c-106">値</span><span class="sxs-lookup"><span data-stu-id="0cb5c-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cb5c-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0cb5c-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cb5c-108">1.0</span><span class="sxs-lookup"><span data-stu-id="0cb5c-108">1.0</span></span>|
|[<span data-ttu-id="0cb5c-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0cb5c-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cb5c-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cb5c-110">ReadItem</span></span>|
|[<span data-ttu-id="0cb5c-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0cb5c-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cb5c-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0cb5c-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="0cb5c-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="0cb5c-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="0cb5c-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="0cb5c-114">displayName :String</span></span>

<span data-ttu-id="0cb5c-115">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="0cb5c-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="0cb5c-116">型</span><span class="sxs-lookup"><span data-stu-id="0cb5c-116">Type</span></span>

*   <span data-ttu-id="0cb5c-117">String</span><span class="sxs-lookup"><span data-stu-id="0cb5c-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cb5c-118">要件</span><span class="sxs-lookup"><span data-stu-id="0cb5c-118">Requirements</span></span>

|<span data-ttu-id="0cb5c-119">要件</span><span class="sxs-lookup"><span data-stu-id="0cb5c-119">Requirement</span></span>| <span data-ttu-id="0cb5c-120">値</span><span class="sxs-lookup"><span data-stu-id="0cb5c-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cb5c-121">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0cb5c-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cb5c-122">1.0</span><span class="sxs-lookup"><span data-stu-id="0cb5c-122">1.0</span></span>|
|[<span data-ttu-id="0cb5c-123">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0cb5c-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cb5c-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cb5c-124">ReadItem</span></span>|
|[<span data-ttu-id="0cb5c-125">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0cb5c-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cb5c-126">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0cb5c-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0cb5c-127">例</span><span class="sxs-lookup"><span data-stu-id="0cb5c-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="0cb5c-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="0cb5c-128">emailAddress :String</span></span>

<span data-ttu-id="0cb5c-129">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="0cb5c-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="0cb5c-130">型</span><span class="sxs-lookup"><span data-stu-id="0cb5c-130">Type</span></span>

*   <span data-ttu-id="0cb5c-131">String</span><span class="sxs-lookup"><span data-stu-id="0cb5c-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cb5c-132">要件</span><span class="sxs-lookup"><span data-stu-id="0cb5c-132">Requirements</span></span>

|<span data-ttu-id="0cb5c-133">要件</span><span class="sxs-lookup"><span data-stu-id="0cb5c-133">Requirement</span></span>| <span data-ttu-id="0cb5c-134">値</span><span class="sxs-lookup"><span data-stu-id="0cb5c-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cb5c-135">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0cb5c-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cb5c-136">1.0</span><span class="sxs-lookup"><span data-stu-id="0cb5c-136">1.0</span></span>|
|[<span data-ttu-id="0cb5c-137">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0cb5c-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cb5c-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cb5c-138">ReadItem</span></span>|
|[<span data-ttu-id="0cb5c-139">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0cb5c-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cb5c-140">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0cb5c-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0cb5c-141">例</span><span class="sxs-lookup"><span data-stu-id="0cb5c-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="0cb5c-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="0cb5c-142">timeZone :String</span></span>

<span data-ttu-id="0cb5c-143">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="0cb5c-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="0cb5c-144">型</span><span class="sxs-lookup"><span data-stu-id="0cb5c-144">Type</span></span>

*   <span data-ttu-id="0cb5c-145">String</span><span class="sxs-lookup"><span data-stu-id="0cb5c-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0cb5c-146">要件</span><span class="sxs-lookup"><span data-stu-id="0cb5c-146">Requirements</span></span>

|<span data-ttu-id="0cb5c-147">要件</span><span class="sxs-lookup"><span data-stu-id="0cb5c-147">Requirement</span></span>| <span data-ttu-id="0cb5c-148">値</span><span class="sxs-lookup"><span data-stu-id="0cb5c-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="0cb5c-149">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0cb5c-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0cb5c-150">1.0</span><span class="sxs-lookup"><span data-stu-id="0cb5c-150">1.0</span></span>|
|[<span data-ttu-id="0cb5c-151">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0cb5c-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0cb5c-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0cb5c-152">ReadItem</span></span>|
|[<span data-ttu-id="0cb5c-153">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0cb5c-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0cb5c-154">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0cb5c-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0cb5c-155">例</span><span class="sxs-lookup"><span data-stu-id="0cb5c-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
