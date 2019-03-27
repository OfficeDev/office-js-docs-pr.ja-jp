---
title: Office.-mailbox-要件セット1.1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 7a10a35887d31a8803d0662eedbe190543d2326a
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870192"
---
# <a name="userprofile"></a><span data-ttu-id="01014-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="01014-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="01014-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="01014-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="01014-104">要件</span><span class="sxs-lookup"><span data-stu-id="01014-104">Requirements</span></span>

|<span data-ttu-id="01014-105">要件</span><span class="sxs-lookup"><span data-stu-id="01014-105">Requirement</span></span>| <span data-ttu-id="01014-106">値</span><span class="sxs-lookup"><span data-stu-id="01014-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="01014-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="01014-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="01014-108">1.0</span><span class="sxs-lookup"><span data-stu-id="01014-108">1.0</span></span>|
|[<span data-ttu-id="01014-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="01014-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="01014-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="01014-110">ReadItem</span></span>|
|[<span data-ttu-id="01014-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="01014-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="01014-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="01014-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="01014-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="01014-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="01014-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="01014-114">displayName :String</span></span>

<span data-ttu-id="01014-115">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="01014-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="01014-116">型</span><span class="sxs-lookup"><span data-stu-id="01014-116">Type</span></span>

*   <span data-ttu-id="01014-117">String</span><span class="sxs-lookup"><span data-stu-id="01014-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="01014-118">要件</span><span class="sxs-lookup"><span data-stu-id="01014-118">Requirements</span></span>

|<span data-ttu-id="01014-119">要件</span><span class="sxs-lookup"><span data-stu-id="01014-119">Requirement</span></span>| <span data-ttu-id="01014-120">値</span><span class="sxs-lookup"><span data-stu-id="01014-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="01014-121">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="01014-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="01014-122">1.0</span><span class="sxs-lookup"><span data-stu-id="01014-122">1.0</span></span>|
|[<span data-ttu-id="01014-123">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="01014-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="01014-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="01014-124">ReadItem</span></span>|
|[<span data-ttu-id="01014-125">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="01014-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="01014-126">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="01014-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="01014-127">例</span><span class="sxs-lookup"><span data-stu-id="01014-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="01014-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="01014-128">emailAddress :String</span></span>

<span data-ttu-id="01014-129">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="01014-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="01014-130">型</span><span class="sxs-lookup"><span data-stu-id="01014-130">Type</span></span>

*   <span data-ttu-id="01014-131">String</span><span class="sxs-lookup"><span data-stu-id="01014-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="01014-132">要件</span><span class="sxs-lookup"><span data-stu-id="01014-132">Requirements</span></span>

|<span data-ttu-id="01014-133">要件</span><span class="sxs-lookup"><span data-stu-id="01014-133">Requirement</span></span>| <span data-ttu-id="01014-134">値</span><span class="sxs-lookup"><span data-stu-id="01014-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="01014-135">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="01014-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="01014-136">1.0</span><span class="sxs-lookup"><span data-stu-id="01014-136">1.0</span></span>|
|[<span data-ttu-id="01014-137">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="01014-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="01014-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="01014-138">ReadItem</span></span>|
|[<span data-ttu-id="01014-139">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="01014-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="01014-140">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="01014-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="01014-141">例</span><span class="sxs-lookup"><span data-stu-id="01014-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="01014-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="01014-142">timeZone :String</span></span>

<span data-ttu-id="01014-143">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="01014-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="01014-144">型</span><span class="sxs-lookup"><span data-stu-id="01014-144">Type</span></span>

*   <span data-ttu-id="01014-145">String</span><span class="sxs-lookup"><span data-stu-id="01014-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="01014-146">要件</span><span class="sxs-lookup"><span data-stu-id="01014-146">Requirements</span></span>

|<span data-ttu-id="01014-147">要件</span><span class="sxs-lookup"><span data-stu-id="01014-147">Requirement</span></span>| <span data-ttu-id="01014-148">値</span><span class="sxs-lookup"><span data-stu-id="01014-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="01014-149">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="01014-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="01014-150">1.0</span><span class="sxs-lookup"><span data-stu-id="01014-150">1.0</span></span>|
|[<span data-ttu-id="01014-151">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="01014-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="01014-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="01014-152">ReadItem</span></span>|
|[<span data-ttu-id="01014-153">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="01014-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="01014-154">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="01014-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="01014-155">例</span><span class="sxs-lookup"><span data-stu-id="01014-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
