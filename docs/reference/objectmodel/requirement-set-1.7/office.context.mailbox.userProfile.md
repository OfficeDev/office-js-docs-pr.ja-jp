---
title: Office.-mailbox-要件セット1.7
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 036f18e4cb98cfe510a19d85a5a79f393ca8bd17
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/26/2019
ms.locfileid: "33353294"
---
# <a name="userprofile"></a><span data-ttu-id="4e0d8-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="4e0d8-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="4e0d8-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="4e0d8-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e0d8-104">要件</span><span class="sxs-lookup"><span data-stu-id="4e0d8-104">Requirements</span></span>

|<span data-ttu-id="4e0d8-105">要件</span><span class="sxs-lookup"><span data-stu-id="4e0d8-105">Requirement</span></span>| <span data-ttu-id="4e0d8-106">値</span><span class="sxs-lookup"><span data-stu-id="4e0d8-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e0d8-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4e0d8-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e0d8-108">1.0</span><span class="sxs-lookup"><span data-stu-id="4e0d8-108">1.0</span></span>|
|[<span data-ttu-id="4e0d8-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4e0d8-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e0d8-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e0d8-110">ReadItem</span></span>|
|[<span data-ttu-id="4e0d8-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4e0d8-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e0d8-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="4e0d8-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4e0d8-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="4e0d8-113">Members and methods</span></span>

| <span data-ttu-id="4e0d8-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="4e0d8-114">Member</span></span> | <span data-ttu-id="4e0d8-115">種類</span><span class="sxs-lookup"><span data-stu-id="4e0d8-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4e0d8-116">accountType</span><span class="sxs-lookup"><span data-stu-id="4e0d8-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="4e0d8-117">Member</span><span class="sxs-lookup"><span data-stu-id="4e0d8-117">Member</span></span> |
| [<span data-ttu-id="4e0d8-118">displayName</span><span class="sxs-lookup"><span data-stu-id="4e0d8-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="4e0d8-119">Member</span><span class="sxs-lookup"><span data-stu-id="4e0d8-119">Member</span></span> |
| [<span data-ttu-id="4e0d8-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="4e0d8-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="4e0d8-121">Member</span><span class="sxs-lookup"><span data-stu-id="4e0d8-121">Member</span></span> |
| [<span data-ttu-id="4e0d8-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="4e0d8-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="4e0d8-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="4e0d8-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="4e0d8-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="4e0d8-124">Members</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="4e0d8-125">accountType: String</span><span class="sxs-lookup"><span data-stu-id="4e0d8-125">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="4e0d8-126">このメンバーは、現在、Outlook 2016 for Mac (ビルド16.9.1212 以降) でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="4e0d8-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="4e0d8-127">メールボックスに関連付けられているユーザーのアカウントの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="4e0d8-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="4e0d8-128">次の表に使用可能な値を示します。</span><span class="sxs-lookup"><span data-stu-id="4e0d8-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="4e0d8-129">値</span><span class="sxs-lookup"><span data-stu-id="4e0d8-129">Value</span></span> | <span data-ttu-id="4e0d8-130">説明</span><span class="sxs-lookup"><span data-stu-id="4e0d8-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="4e0d8-131">メールボックスは、オンプレミスの Exchange サーバーにあります。</span><span class="sxs-lookup"><span data-stu-id="4e0d8-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="4e0d8-132">メールボックスは、Gmail アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="4e0d8-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="4e0d8-133">メールボックスは、Office 365 の職場または学校のアカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="4e0d8-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="4e0d8-134">メールボックスは、個人の Outlook.com アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="4e0d8-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="4e0d8-135">型</span><span class="sxs-lookup"><span data-stu-id="4e0d8-135">Type</span></span>

*   <span data-ttu-id="4e0d8-136">String</span><span class="sxs-lookup"><span data-stu-id="4e0d8-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e0d8-137">要件</span><span class="sxs-lookup"><span data-stu-id="4e0d8-137">Requirements</span></span>

|<span data-ttu-id="4e0d8-138">要件</span><span class="sxs-lookup"><span data-stu-id="4e0d8-138">Requirement</span></span>| <span data-ttu-id="4e0d8-139">値</span><span class="sxs-lookup"><span data-stu-id="4e0d8-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e0d8-140">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4e0d8-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e0d8-141">1.6</span><span class="sxs-lookup"><span data-stu-id="4e0d8-141">1.6</span></span> |
|[<span data-ttu-id="4e0d8-142">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4e0d8-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e0d8-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e0d8-143">ReadItem</span></span>|
|[<span data-ttu-id="4e0d8-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4e0d8-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e0d8-145">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="4e0d8-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e0d8-146">例</span><span class="sxs-lookup"><span data-stu-id="4e0d8-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

---
---

#### <a name="displayname-string"></a><span data-ttu-id="4e0d8-147">displayName: String</span><span class="sxs-lookup"><span data-stu-id="4e0d8-147">displayName: String</span></span>

<span data-ttu-id="4e0d8-148">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="4e0d8-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="4e0d8-149">型</span><span class="sxs-lookup"><span data-stu-id="4e0d8-149">Type</span></span>

*   <span data-ttu-id="4e0d8-150">String</span><span class="sxs-lookup"><span data-stu-id="4e0d8-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e0d8-151">要件</span><span class="sxs-lookup"><span data-stu-id="4e0d8-151">Requirements</span></span>

|<span data-ttu-id="4e0d8-152">要件</span><span class="sxs-lookup"><span data-stu-id="4e0d8-152">Requirement</span></span>| <span data-ttu-id="4e0d8-153">値</span><span class="sxs-lookup"><span data-stu-id="4e0d8-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e0d8-154">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4e0d8-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e0d8-155">1.0</span><span class="sxs-lookup"><span data-stu-id="4e0d8-155">1.0</span></span>|
|[<span data-ttu-id="4e0d8-156">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4e0d8-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e0d8-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e0d8-157">ReadItem</span></span>|
|[<span data-ttu-id="4e0d8-158">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4e0d8-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e0d8-159">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="4e0d8-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e0d8-160">例</span><span class="sxs-lookup"><span data-stu-id="4e0d8-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="4e0d8-161">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="4e0d8-161">emailAddress: String</span></span>

<span data-ttu-id="4e0d8-162">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="4e0d8-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="4e0d8-163">型</span><span class="sxs-lookup"><span data-stu-id="4e0d8-163">Type</span></span>

*   <span data-ttu-id="4e0d8-164">String</span><span class="sxs-lookup"><span data-stu-id="4e0d8-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e0d8-165">要件</span><span class="sxs-lookup"><span data-stu-id="4e0d8-165">Requirements</span></span>

|<span data-ttu-id="4e0d8-166">要件</span><span class="sxs-lookup"><span data-stu-id="4e0d8-166">Requirement</span></span>| <span data-ttu-id="4e0d8-167">値</span><span class="sxs-lookup"><span data-stu-id="4e0d8-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e0d8-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4e0d8-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e0d8-169">1.0</span><span class="sxs-lookup"><span data-stu-id="4e0d8-169">1.0</span></span>|
|[<span data-ttu-id="4e0d8-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4e0d8-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e0d8-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e0d8-171">ReadItem</span></span>|
|[<span data-ttu-id="4e0d8-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4e0d8-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e0d8-173">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="4e0d8-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e0d8-174">例</span><span class="sxs-lookup"><span data-stu-id="4e0d8-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

---
---

#### <a name="timezone-string"></a><span data-ttu-id="4e0d8-175">timeZone: String</span><span class="sxs-lookup"><span data-stu-id="4e0d8-175">timeZone: String</span></span>

<span data-ttu-id="4e0d8-176">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="4e0d8-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="4e0d8-177">型</span><span class="sxs-lookup"><span data-stu-id="4e0d8-177">Type</span></span>

*   <span data-ttu-id="4e0d8-178">String</span><span class="sxs-lookup"><span data-stu-id="4e0d8-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e0d8-179">要件</span><span class="sxs-lookup"><span data-stu-id="4e0d8-179">Requirements</span></span>

|<span data-ttu-id="4e0d8-180">要件</span><span class="sxs-lookup"><span data-stu-id="4e0d8-180">Requirement</span></span>| <span data-ttu-id="4e0d8-181">値</span><span class="sxs-lookup"><span data-stu-id="4e0d8-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e0d8-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4e0d8-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e0d8-183">1.0</span><span class="sxs-lookup"><span data-stu-id="4e0d8-183">1.0</span></span>|
|[<span data-ttu-id="4e0d8-184">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4e0d8-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e0d8-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e0d8-185">ReadItem</span></span>|
|[<span data-ttu-id="4e0d8-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4e0d8-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e0d8-187">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="4e0d8-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e0d8-188">例</span><span class="sxs-lookup"><span data-stu-id="4e0d8-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
