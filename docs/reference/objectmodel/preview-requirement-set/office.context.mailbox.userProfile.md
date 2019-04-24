---
title: Office.context.mailbox.userProfile - プレビュー要件セット
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 29111314f16bb9c6518b350254a3036ffa125796
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451963"
---
# <a name="userprofile"></a><span data-ttu-id="2c885-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="2c885-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="2c885-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="2c885-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="2c885-104">要件</span><span class="sxs-lookup"><span data-stu-id="2c885-104">Requirements</span></span>

|<span data-ttu-id="2c885-105">要件</span><span class="sxs-lookup"><span data-stu-id="2c885-105">Requirement</span></span>| <span data-ttu-id="2c885-106">値</span><span class="sxs-lookup"><span data-stu-id="2c885-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c885-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2c885-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2c885-108">1.0</span><span class="sxs-lookup"><span data-stu-id="2c885-108">1.0</span></span>|
|[<span data-ttu-id="2c885-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2c885-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2c885-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2c885-110">ReadItem</span></span>|
|[<span data-ttu-id="2c885-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2c885-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2c885-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2c885-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="2c885-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="2c885-113">Members and methods</span></span>

| <span data-ttu-id="2c885-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="2c885-114">Member</span></span> | <span data-ttu-id="2c885-115">種類</span><span class="sxs-lookup"><span data-stu-id="2c885-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="2c885-116">accountType</span><span class="sxs-lookup"><span data-stu-id="2c885-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="2c885-117">Member</span><span class="sxs-lookup"><span data-stu-id="2c885-117">Member</span></span> |
| [<span data-ttu-id="2c885-118">displayName</span><span class="sxs-lookup"><span data-stu-id="2c885-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="2c885-119">Member</span><span class="sxs-lookup"><span data-stu-id="2c885-119">Member</span></span> |
| [<span data-ttu-id="2c885-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="2c885-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="2c885-121">Member</span><span class="sxs-lookup"><span data-stu-id="2c885-121">Member</span></span> |
| [<span data-ttu-id="2c885-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="2c885-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="2c885-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="2c885-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="2c885-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="2c885-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="2c885-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="2c885-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="2c885-126">現在、このメンバーは Outlook 2016 for Mac 以降 (ビルド 16.9.1212 以降) でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="2c885-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="2c885-127">メールボックスに関連付けられているユーザーのアカウントの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="2c885-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="2c885-128">次の表に使用可能な値を示します。</span><span class="sxs-lookup"><span data-stu-id="2c885-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="2c885-129">値</span><span class="sxs-lookup"><span data-stu-id="2c885-129">Value</span></span> | <span data-ttu-id="2c885-130">説明</span><span class="sxs-lookup"><span data-stu-id="2c885-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="2c885-131">メールボックスは、オンプレミスの Exchange サーバーにあります。</span><span class="sxs-lookup"><span data-stu-id="2c885-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="2c885-132">メールボックスは、Gmail アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="2c885-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="2c885-133">メールボックスは、Office 365 の職場または学校のアカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="2c885-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="2c885-134">メールボックスは、個人の Outlook.com アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="2c885-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="2c885-135">型</span><span class="sxs-lookup"><span data-stu-id="2c885-135">Type</span></span>

*   <span data-ttu-id="2c885-136">String</span><span class="sxs-lookup"><span data-stu-id="2c885-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2c885-137">要件</span><span class="sxs-lookup"><span data-stu-id="2c885-137">Requirements</span></span>

|<span data-ttu-id="2c885-138">要件</span><span class="sxs-lookup"><span data-stu-id="2c885-138">Requirement</span></span>| <span data-ttu-id="2c885-139">値</span><span class="sxs-lookup"><span data-stu-id="2c885-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c885-140">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2c885-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2c885-141">1.6</span><span class="sxs-lookup"><span data-stu-id="2c885-141">1.6</span></span> |
|[<span data-ttu-id="2c885-142">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2c885-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2c885-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2c885-143">ReadItem</span></span>|
|[<span data-ttu-id="2c885-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2c885-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2c885-145">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2c885-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2c885-146">例</span><span class="sxs-lookup"><span data-stu-id="2c885-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

---
---

####  <a name="displayname-string"></a><span data-ttu-id="2c885-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="2c885-147">displayName :String</span></span>

<span data-ttu-id="2c885-148">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="2c885-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="2c885-149">型</span><span class="sxs-lookup"><span data-stu-id="2c885-149">Type</span></span>

*   <span data-ttu-id="2c885-150">String</span><span class="sxs-lookup"><span data-stu-id="2c885-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2c885-151">要件</span><span class="sxs-lookup"><span data-stu-id="2c885-151">Requirements</span></span>

|<span data-ttu-id="2c885-152">要件</span><span class="sxs-lookup"><span data-stu-id="2c885-152">Requirement</span></span>| <span data-ttu-id="2c885-153">値</span><span class="sxs-lookup"><span data-stu-id="2c885-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c885-154">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2c885-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2c885-155">1.0</span><span class="sxs-lookup"><span data-stu-id="2c885-155">1.0</span></span>|
|[<span data-ttu-id="2c885-156">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2c885-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2c885-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2c885-157">ReadItem</span></span>|
|[<span data-ttu-id="2c885-158">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2c885-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2c885-159">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2c885-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2c885-160">例</span><span class="sxs-lookup"><span data-stu-id="2c885-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

---
---

####  <a name="emailaddress-string"></a><span data-ttu-id="2c885-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="2c885-161">emailAddress :String</span></span>

<span data-ttu-id="2c885-162">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="2c885-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="2c885-163">型</span><span class="sxs-lookup"><span data-stu-id="2c885-163">Type</span></span>

*   <span data-ttu-id="2c885-164">String</span><span class="sxs-lookup"><span data-stu-id="2c885-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2c885-165">要件</span><span class="sxs-lookup"><span data-stu-id="2c885-165">Requirements</span></span>

|<span data-ttu-id="2c885-166">要件</span><span class="sxs-lookup"><span data-stu-id="2c885-166">Requirement</span></span>| <span data-ttu-id="2c885-167">値</span><span class="sxs-lookup"><span data-stu-id="2c885-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c885-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2c885-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2c885-169">1.0</span><span class="sxs-lookup"><span data-stu-id="2c885-169">1.0</span></span>|
|[<span data-ttu-id="2c885-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2c885-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2c885-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2c885-171">ReadItem</span></span>|
|[<span data-ttu-id="2c885-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2c885-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2c885-173">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2c885-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2c885-174">例</span><span class="sxs-lookup"><span data-stu-id="2c885-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

---
---

####  <a name="timezone-string"></a><span data-ttu-id="2c885-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="2c885-175">timeZone :String</span></span>

<span data-ttu-id="2c885-176">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="2c885-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="2c885-177">型</span><span class="sxs-lookup"><span data-stu-id="2c885-177">Type</span></span>

*   <span data-ttu-id="2c885-178">String</span><span class="sxs-lookup"><span data-stu-id="2c885-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2c885-179">要件</span><span class="sxs-lookup"><span data-stu-id="2c885-179">Requirements</span></span>

|<span data-ttu-id="2c885-180">要件</span><span class="sxs-lookup"><span data-stu-id="2c885-180">Requirement</span></span>| <span data-ttu-id="2c885-181">値</span><span class="sxs-lookup"><span data-stu-id="2c885-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c885-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2c885-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2c885-183">1.0</span><span class="sxs-lookup"><span data-stu-id="2c885-183">1.0</span></span>|
|[<span data-ttu-id="2c885-184">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2c885-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2c885-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2c885-185">ReadItem</span></span>|
|[<span data-ttu-id="2c885-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2c885-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2c885-187">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2c885-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2c885-188">例</span><span class="sxs-lookup"><span data-stu-id="2c885-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
