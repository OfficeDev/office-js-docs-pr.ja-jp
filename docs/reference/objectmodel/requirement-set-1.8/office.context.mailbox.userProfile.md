---
title: Office.-mailbox-要件セット1.8
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: 39a833a81eab22c70d89cdfc61784555312b23d6
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902196"
---
# <a name="userprofile"></a><span data-ttu-id="493c6-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="493c6-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="493c6-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="493c6-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="493c6-104">要件</span><span class="sxs-lookup"><span data-stu-id="493c6-104">Requirements</span></span>

|<span data-ttu-id="493c6-105">要件</span><span class="sxs-lookup"><span data-stu-id="493c6-105">Requirement</span></span>| <span data-ttu-id="493c6-106">値</span><span class="sxs-lookup"><span data-stu-id="493c6-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="493c6-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="493c6-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="493c6-108">1.0</span><span class="sxs-lookup"><span data-stu-id="493c6-108">1.0</span></span>|
|[<span data-ttu-id="493c6-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="493c6-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="493c6-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="493c6-110">ReadItem</span></span>|
|[<span data-ttu-id="493c6-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="493c6-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="493c6-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="493c6-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="493c6-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="493c6-113">Members and methods</span></span>

| <span data-ttu-id="493c6-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="493c6-114">Member</span></span> | <span data-ttu-id="493c6-115">種類</span><span class="sxs-lookup"><span data-stu-id="493c6-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="493c6-116">accountType</span><span class="sxs-lookup"><span data-stu-id="493c6-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="493c6-117">Member</span><span class="sxs-lookup"><span data-stu-id="493c6-117">Member</span></span> |
| [<span data-ttu-id="493c6-118">displayName</span><span class="sxs-lookup"><span data-stu-id="493c6-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="493c6-119">Member</span><span class="sxs-lookup"><span data-stu-id="493c6-119">Member</span></span> |
| [<span data-ttu-id="493c6-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="493c6-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="493c6-121">Member</span><span class="sxs-lookup"><span data-stu-id="493c6-121">Member</span></span> |
| [<span data-ttu-id="493c6-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="493c6-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="493c6-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="493c6-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="493c6-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="493c6-124">Members</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="493c6-125">accountType: String</span><span class="sxs-lookup"><span data-stu-id="493c6-125">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="493c6-126">このメンバーは、現在、Outlook 2016 以降の Mac (ビルド16.9.1212 以降) でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="493c6-126">This member is currently only supported in Outlook 2016 or later on Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="493c6-127">メールボックスに関連付けられているユーザーのアカウントの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="493c6-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="493c6-128">次の表に使用可能な値を示します。</span><span class="sxs-lookup"><span data-stu-id="493c6-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="493c6-129">値</span><span class="sxs-lookup"><span data-stu-id="493c6-129">Value</span></span> | <span data-ttu-id="493c6-130">説明</span><span class="sxs-lookup"><span data-stu-id="493c6-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="493c6-131">メールボックスは、オンプレミスの Exchange サーバーにあります。</span><span class="sxs-lookup"><span data-stu-id="493c6-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="493c6-132">メールボックスは、Gmail アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="493c6-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="493c6-133">メールボックスは、Office 365 の職場または学校のアカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="493c6-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="493c6-134">メールボックスは、個人の Outlook.com アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="493c6-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="493c6-135">型</span><span class="sxs-lookup"><span data-stu-id="493c6-135">Type</span></span>

*   <span data-ttu-id="493c6-136">String</span><span class="sxs-lookup"><span data-stu-id="493c6-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="493c6-137">要件</span><span class="sxs-lookup"><span data-stu-id="493c6-137">Requirements</span></span>

|<span data-ttu-id="493c6-138">要件</span><span class="sxs-lookup"><span data-stu-id="493c6-138">Requirement</span></span>| <span data-ttu-id="493c6-139">値</span><span class="sxs-lookup"><span data-stu-id="493c6-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="493c6-140">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="493c6-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="493c6-141">1.6</span><span class="sxs-lookup"><span data-stu-id="493c6-141">1.6</span></span> |
|[<span data-ttu-id="493c6-142">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="493c6-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="493c6-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="493c6-143">ReadItem</span></span>|
|[<span data-ttu-id="493c6-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="493c6-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="493c6-145">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="493c6-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="493c6-146">例</span><span class="sxs-lookup"><span data-stu-id="493c6-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

<br>

---
---

#### <a name="displayname-string"></a><span data-ttu-id="493c6-147">displayName: String</span><span class="sxs-lookup"><span data-stu-id="493c6-147">displayName: String</span></span>

<span data-ttu-id="493c6-148">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="493c6-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="493c6-149">型</span><span class="sxs-lookup"><span data-stu-id="493c6-149">Type</span></span>

*   <span data-ttu-id="493c6-150">String</span><span class="sxs-lookup"><span data-stu-id="493c6-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="493c6-151">要件</span><span class="sxs-lookup"><span data-stu-id="493c6-151">Requirements</span></span>

|<span data-ttu-id="493c6-152">要件</span><span class="sxs-lookup"><span data-stu-id="493c6-152">Requirement</span></span>| <span data-ttu-id="493c6-153">値</span><span class="sxs-lookup"><span data-stu-id="493c6-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="493c6-154">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="493c6-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="493c6-155">1.0</span><span class="sxs-lookup"><span data-stu-id="493c6-155">1.0</span></span>|
|[<span data-ttu-id="493c6-156">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="493c6-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="493c6-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="493c6-157">ReadItem</span></span>|
|[<span data-ttu-id="493c6-158">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="493c6-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="493c6-159">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="493c6-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="493c6-160">例</span><span class="sxs-lookup"><span data-stu-id="493c6-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="493c6-161">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="493c6-161">emailAddress: String</span></span>

<span data-ttu-id="493c6-162">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="493c6-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="493c6-163">型</span><span class="sxs-lookup"><span data-stu-id="493c6-163">Type</span></span>

*   <span data-ttu-id="493c6-164">String</span><span class="sxs-lookup"><span data-stu-id="493c6-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="493c6-165">要件</span><span class="sxs-lookup"><span data-stu-id="493c6-165">Requirements</span></span>

|<span data-ttu-id="493c6-166">要件</span><span class="sxs-lookup"><span data-stu-id="493c6-166">Requirement</span></span>| <span data-ttu-id="493c6-167">値</span><span class="sxs-lookup"><span data-stu-id="493c6-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="493c6-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="493c6-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="493c6-169">1.0</span><span class="sxs-lookup"><span data-stu-id="493c6-169">1.0</span></span>|
|[<span data-ttu-id="493c6-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="493c6-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="493c6-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="493c6-171">ReadItem</span></span>|
|[<span data-ttu-id="493c6-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="493c6-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="493c6-173">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="493c6-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="493c6-174">例</span><span class="sxs-lookup"><span data-stu-id="493c6-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="493c6-175">timeZone: String</span><span class="sxs-lookup"><span data-stu-id="493c6-175">timeZone: String</span></span>

<span data-ttu-id="493c6-176">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="493c6-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="493c6-177">型</span><span class="sxs-lookup"><span data-stu-id="493c6-177">Type</span></span>

*   <span data-ttu-id="493c6-178">String</span><span class="sxs-lookup"><span data-stu-id="493c6-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="493c6-179">要件</span><span class="sxs-lookup"><span data-stu-id="493c6-179">Requirements</span></span>

|<span data-ttu-id="493c6-180">要件</span><span class="sxs-lookup"><span data-stu-id="493c6-180">Requirement</span></span>| <span data-ttu-id="493c6-181">値</span><span class="sxs-lookup"><span data-stu-id="493c6-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="493c6-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="493c6-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="493c6-183">1.0</span><span class="sxs-lookup"><span data-stu-id="493c6-183">1.0</span></span>|
|[<span data-ttu-id="493c6-184">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="493c6-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="493c6-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="493c6-185">ReadItem</span></span>|
|[<span data-ttu-id="493c6-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="493c6-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="493c6-187">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="493c6-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="493c6-188">例</span><span class="sxs-lookup"><span data-stu-id="493c6-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
