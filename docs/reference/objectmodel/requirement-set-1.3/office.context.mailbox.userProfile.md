---
title: Office.-mailbox-要件セット1.3
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 03cdc13845bff0fbd3855f29f43298cd770e5ad9
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30869912"
---
# <a name="userprofile"></a><span data-ttu-id="37904-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="37904-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="37904-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="37904-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="37904-104">要件</span><span class="sxs-lookup"><span data-stu-id="37904-104">Requirements</span></span>

|<span data-ttu-id="37904-105">要件</span><span class="sxs-lookup"><span data-stu-id="37904-105">Requirement</span></span>| <span data-ttu-id="37904-106">値</span><span class="sxs-lookup"><span data-stu-id="37904-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="37904-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="37904-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37904-108">1.0</span><span class="sxs-lookup"><span data-stu-id="37904-108">1.0</span></span>|
|[<span data-ttu-id="37904-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="37904-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37904-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37904-110">ReadItem</span></span>|
|[<span data-ttu-id="37904-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="37904-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="37904-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="37904-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="37904-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="37904-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="37904-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="37904-114">displayName :String</span></span>

<span data-ttu-id="37904-115">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="37904-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="37904-116">型</span><span class="sxs-lookup"><span data-stu-id="37904-116">Type</span></span>

*   <span data-ttu-id="37904-117">String</span><span class="sxs-lookup"><span data-stu-id="37904-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="37904-118">要件</span><span class="sxs-lookup"><span data-stu-id="37904-118">Requirements</span></span>

|<span data-ttu-id="37904-119">要件</span><span class="sxs-lookup"><span data-stu-id="37904-119">Requirement</span></span>| <span data-ttu-id="37904-120">値</span><span class="sxs-lookup"><span data-stu-id="37904-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="37904-121">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="37904-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37904-122">1.0</span><span class="sxs-lookup"><span data-stu-id="37904-122">1.0</span></span>|
|[<span data-ttu-id="37904-123">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="37904-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37904-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37904-124">ReadItem</span></span>|
|[<span data-ttu-id="37904-125">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="37904-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="37904-126">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="37904-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37904-127">例</span><span class="sxs-lookup"><span data-stu-id="37904-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="37904-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="37904-128">emailAddress :String</span></span>

<span data-ttu-id="37904-129">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="37904-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="37904-130">型</span><span class="sxs-lookup"><span data-stu-id="37904-130">Type</span></span>

*   <span data-ttu-id="37904-131">String</span><span class="sxs-lookup"><span data-stu-id="37904-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="37904-132">要件</span><span class="sxs-lookup"><span data-stu-id="37904-132">Requirements</span></span>

|<span data-ttu-id="37904-133">要件</span><span class="sxs-lookup"><span data-stu-id="37904-133">Requirement</span></span>| <span data-ttu-id="37904-134">値</span><span class="sxs-lookup"><span data-stu-id="37904-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="37904-135">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="37904-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37904-136">1.0</span><span class="sxs-lookup"><span data-stu-id="37904-136">1.0</span></span>|
|[<span data-ttu-id="37904-137">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="37904-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37904-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37904-138">ReadItem</span></span>|
|[<span data-ttu-id="37904-139">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="37904-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="37904-140">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="37904-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37904-141">例</span><span class="sxs-lookup"><span data-stu-id="37904-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="37904-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="37904-142">timeZone :String</span></span>

<span data-ttu-id="37904-143">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="37904-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="37904-144">型</span><span class="sxs-lookup"><span data-stu-id="37904-144">Type</span></span>

*   <span data-ttu-id="37904-145">String</span><span class="sxs-lookup"><span data-stu-id="37904-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="37904-146">要件</span><span class="sxs-lookup"><span data-stu-id="37904-146">Requirements</span></span>

|<span data-ttu-id="37904-147">要件</span><span class="sxs-lookup"><span data-stu-id="37904-147">Requirement</span></span>| <span data-ttu-id="37904-148">値</span><span class="sxs-lookup"><span data-stu-id="37904-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="37904-149">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="37904-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="37904-150">1.0</span><span class="sxs-lookup"><span data-stu-id="37904-150">1.0</span></span>|
|[<span data-ttu-id="37904-151">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="37904-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="37904-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="37904-152">ReadItem</span></span>|
|[<span data-ttu-id="37904-153">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="37904-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="37904-154">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="37904-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="37904-155">例</span><span class="sxs-lookup"><span data-stu-id="37904-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
