---
title: Office.-mailbox-要件セット1.2
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 496a59f4ef02f03cda95fde0bf14634b1db13f77
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870584"
---
# <a name="userprofile"></a><span data-ttu-id="a1eb4-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="a1eb4-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="a1eb4-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="a1eb4-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1eb4-104">要件</span><span class="sxs-lookup"><span data-stu-id="a1eb4-104">Requirements</span></span>

|<span data-ttu-id="a1eb4-105">要件</span><span class="sxs-lookup"><span data-stu-id="a1eb4-105">Requirement</span></span>| <span data-ttu-id="a1eb4-106">値</span><span class="sxs-lookup"><span data-stu-id="a1eb4-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1eb4-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a1eb4-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1eb4-108">1.0</span><span class="sxs-lookup"><span data-stu-id="a1eb4-108">1.0</span></span>|
|[<span data-ttu-id="a1eb4-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a1eb4-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1eb4-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1eb4-110">ReadItem</span></span>|
|[<span data-ttu-id="a1eb4-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a1eb4-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a1eb4-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="a1eb4-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="a1eb4-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="a1eb4-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="a1eb4-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="a1eb4-114">displayName :String</span></span>

<span data-ttu-id="a1eb4-115">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="a1eb4-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="a1eb4-116">型</span><span class="sxs-lookup"><span data-stu-id="a1eb4-116">Type</span></span>

*   <span data-ttu-id="a1eb4-117">String</span><span class="sxs-lookup"><span data-stu-id="a1eb4-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1eb4-118">要件</span><span class="sxs-lookup"><span data-stu-id="a1eb4-118">Requirements</span></span>

|<span data-ttu-id="a1eb4-119">要件</span><span class="sxs-lookup"><span data-stu-id="a1eb4-119">Requirement</span></span>| <span data-ttu-id="a1eb4-120">値</span><span class="sxs-lookup"><span data-stu-id="a1eb4-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1eb4-121">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a1eb4-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1eb4-122">1.0</span><span class="sxs-lookup"><span data-stu-id="a1eb4-122">1.0</span></span>|
|[<span data-ttu-id="a1eb4-123">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a1eb4-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1eb4-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1eb4-124">ReadItem</span></span>|
|[<span data-ttu-id="a1eb4-125">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a1eb4-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a1eb4-126">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="a1eb4-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1eb4-127">例</span><span class="sxs-lookup"><span data-stu-id="a1eb4-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="a1eb4-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="a1eb4-128">emailAddress :String</span></span>

<span data-ttu-id="a1eb4-129">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="a1eb4-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="a1eb4-130">型</span><span class="sxs-lookup"><span data-stu-id="a1eb4-130">Type</span></span>

*   <span data-ttu-id="a1eb4-131">String</span><span class="sxs-lookup"><span data-stu-id="a1eb4-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1eb4-132">要件</span><span class="sxs-lookup"><span data-stu-id="a1eb4-132">Requirements</span></span>

|<span data-ttu-id="a1eb4-133">要件</span><span class="sxs-lookup"><span data-stu-id="a1eb4-133">Requirement</span></span>| <span data-ttu-id="a1eb4-134">値</span><span class="sxs-lookup"><span data-stu-id="a1eb4-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1eb4-135">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a1eb4-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1eb4-136">1.0</span><span class="sxs-lookup"><span data-stu-id="a1eb4-136">1.0</span></span>|
|[<span data-ttu-id="a1eb4-137">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a1eb4-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1eb4-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1eb4-138">ReadItem</span></span>|
|[<span data-ttu-id="a1eb4-139">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a1eb4-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a1eb4-140">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="a1eb4-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1eb4-141">例</span><span class="sxs-lookup"><span data-stu-id="a1eb4-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="a1eb4-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="a1eb4-142">timeZone :String</span></span>

<span data-ttu-id="a1eb4-143">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="a1eb4-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="a1eb4-144">型</span><span class="sxs-lookup"><span data-stu-id="a1eb4-144">Type</span></span>

*   <span data-ttu-id="a1eb4-145">String</span><span class="sxs-lookup"><span data-stu-id="a1eb4-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1eb4-146">要件</span><span class="sxs-lookup"><span data-stu-id="a1eb4-146">Requirements</span></span>

|<span data-ttu-id="a1eb4-147">要件</span><span class="sxs-lookup"><span data-stu-id="a1eb4-147">Requirement</span></span>| <span data-ttu-id="a1eb4-148">値</span><span class="sxs-lookup"><span data-stu-id="a1eb4-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1eb4-149">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a1eb4-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1eb4-150">1.0</span><span class="sxs-lookup"><span data-stu-id="a1eb4-150">1.0</span></span>|
|[<span data-ttu-id="a1eb4-151">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a1eb4-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1eb4-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a1eb4-152">ReadItem</span></span>|
|[<span data-ttu-id="a1eb4-153">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a1eb4-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a1eb4-154">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="a1eb4-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1eb4-155">例</span><span class="sxs-lookup"><span data-stu-id="a1eb4-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
