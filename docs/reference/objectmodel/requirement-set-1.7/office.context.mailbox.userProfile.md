---
title: Office.-mailbox-要件セット1.7
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 8e33d40bec9b561c642ad6e0da73ae13a18378b6
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695904"
---
# <a name="userprofile"></a><span data-ttu-id="28e09-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="28e09-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="28e09-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="28e09-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="28e09-104">要件</span><span class="sxs-lookup"><span data-stu-id="28e09-104">Requirements</span></span>

|<span data-ttu-id="28e09-105">要件</span><span class="sxs-lookup"><span data-stu-id="28e09-105">Requirement</span></span>| <span data-ttu-id="28e09-106">値</span><span class="sxs-lookup"><span data-stu-id="28e09-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="28e09-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="28e09-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="28e09-108">1.0</span><span class="sxs-lookup"><span data-stu-id="28e09-108">1.0</span></span>|
|[<span data-ttu-id="28e09-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="28e09-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="28e09-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28e09-110">ReadItem</span></span>|
|[<span data-ttu-id="28e09-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="28e09-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="28e09-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="28e09-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="28e09-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="28e09-113">Members and methods</span></span>

| <span data-ttu-id="28e09-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="28e09-114">Member</span></span> | <span data-ttu-id="28e09-115">種類</span><span class="sxs-lookup"><span data-stu-id="28e09-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="28e09-116">accountType</span><span class="sxs-lookup"><span data-stu-id="28e09-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="28e09-117">Member</span><span class="sxs-lookup"><span data-stu-id="28e09-117">Member</span></span> |
| [<span data-ttu-id="28e09-118">displayName</span><span class="sxs-lookup"><span data-stu-id="28e09-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="28e09-119">Member</span><span class="sxs-lookup"><span data-stu-id="28e09-119">Member</span></span> |
| [<span data-ttu-id="28e09-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="28e09-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="28e09-121">Member</span><span class="sxs-lookup"><span data-stu-id="28e09-121">Member</span></span> |
| [<span data-ttu-id="28e09-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="28e09-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="28e09-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="28e09-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="28e09-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="28e09-124">Members</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="28e09-125">accountType: String</span><span class="sxs-lookup"><span data-stu-id="28e09-125">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="28e09-126">このメンバーは、現在、Outlook 2016 以降の Mac (ビルド16.9.1212 以降) でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="28e09-126">This member is currently only supported by Outlook 2016 or later on Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="28e09-127">メールボックスに関連付けられているユーザーのアカウントの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="28e09-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="28e09-128">次の表に使用可能な値を示します。</span><span class="sxs-lookup"><span data-stu-id="28e09-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="28e09-129">値</span><span class="sxs-lookup"><span data-stu-id="28e09-129">Value</span></span> | <span data-ttu-id="28e09-130">説明</span><span class="sxs-lookup"><span data-stu-id="28e09-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="28e09-131">メールボックスは、オンプレミスの Exchange サーバーにあります。</span><span class="sxs-lookup"><span data-stu-id="28e09-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="28e09-132">メールボックスは、Gmail アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="28e09-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="28e09-133">メールボックスは、Office 365 の職場または学校のアカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="28e09-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="28e09-134">メールボックスは、個人の Outlook.com アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="28e09-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="28e09-135">型</span><span class="sxs-lookup"><span data-stu-id="28e09-135">Type</span></span>

*   <span data-ttu-id="28e09-136">String</span><span class="sxs-lookup"><span data-stu-id="28e09-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="28e09-137">要件</span><span class="sxs-lookup"><span data-stu-id="28e09-137">Requirements</span></span>

|<span data-ttu-id="28e09-138">要件</span><span class="sxs-lookup"><span data-stu-id="28e09-138">Requirement</span></span>| <span data-ttu-id="28e09-139">値</span><span class="sxs-lookup"><span data-stu-id="28e09-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="28e09-140">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="28e09-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="28e09-141">1.6</span><span class="sxs-lookup"><span data-stu-id="28e09-141">1.6</span></span> |
|[<span data-ttu-id="28e09-142">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="28e09-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="28e09-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28e09-143">ReadItem</span></span>|
|[<span data-ttu-id="28e09-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="28e09-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="28e09-145">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="28e09-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28e09-146">例</span><span class="sxs-lookup"><span data-stu-id="28e09-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

<br>

---
---

#### <a name="displayname-string"></a><span data-ttu-id="28e09-147">displayName: String</span><span class="sxs-lookup"><span data-stu-id="28e09-147">displayName: String</span></span>

<span data-ttu-id="28e09-148">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="28e09-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="28e09-149">型</span><span class="sxs-lookup"><span data-stu-id="28e09-149">Type</span></span>

*   <span data-ttu-id="28e09-150">String</span><span class="sxs-lookup"><span data-stu-id="28e09-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="28e09-151">要件</span><span class="sxs-lookup"><span data-stu-id="28e09-151">Requirements</span></span>

|<span data-ttu-id="28e09-152">要件</span><span class="sxs-lookup"><span data-stu-id="28e09-152">Requirement</span></span>| <span data-ttu-id="28e09-153">値</span><span class="sxs-lookup"><span data-stu-id="28e09-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="28e09-154">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="28e09-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="28e09-155">1.0</span><span class="sxs-lookup"><span data-stu-id="28e09-155">1.0</span></span>|
|[<span data-ttu-id="28e09-156">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="28e09-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="28e09-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28e09-157">ReadItem</span></span>|
|[<span data-ttu-id="28e09-158">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="28e09-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="28e09-159">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="28e09-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28e09-160">例</span><span class="sxs-lookup"><span data-stu-id="28e09-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="28e09-161">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="28e09-161">emailAddress: String</span></span>

<span data-ttu-id="28e09-162">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="28e09-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="28e09-163">型</span><span class="sxs-lookup"><span data-stu-id="28e09-163">Type</span></span>

*   <span data-ttu-id="28e09-164">String</span><span class="sxs-lookup"><span data-stu-id="28e09-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="28e09-165">要件</span><span class="sxs-lookup"><span data-stu-id="28e09-165">Requirements</span></span>

|<span data-ttu-id="28e09-166">要件</span><span class="sxs-lookup"><span data-stu-id="28e09-166">Requirement</span></span>| <span data-ttu-id="28e09-167">値</span><span class="sxs-lookup"><span data-stu-id="28e09-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="28e09-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="28e09-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="28e09-169">1.0</span><span class="sxs-lookup"><span data-stu-id="28e09-169">1.0</span></span>|
|[<span data-ttu-id="28e09-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="28e09-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="28e09-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28e09-171">ReadItem</span></span>|
|[<span data-ttu-id="28e09-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="28e09-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="28e09-173">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="28e09-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28e09-174">例</span><span class="sxs-lookup"><span data-stu-id="28e09-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="28e09-175">timeZone: String</span><span class="sxs-lookup"><span data-stu-id="28e09-175">timeZone: String</span></span>

<span data-ttu-id="28e09-176">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="28e09-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="28e09-177">型</span><span class="sxs-lookup"><span data-stu-id="28e09-177">Type</span></span>

*   <span data-ttu-id="28e09-178">String</span><span class="sxs-lookup"><span data-stu-id="28e09-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="28e09-179">要件</span><span class="sxs-lookup"><span data-stu-id="28e09-179">Requirements</span></span>

|<span data-ttu-id="28e09-180">要件</span><span class="sxs-lookup"><span data-stu-id="28e09-180">Requirement</span></span>| <span data-ttu-id="28e09-181">値</span><span class="sxs-lookup"><span data-stu-id="28e09-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="28e09-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="28e09-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="28e09-183">1.0</span><span class="sxs-lookup"><span data-stu-id="28e09-183">1.0</span></span>|
|[<span data-ttu-id="28e09-184">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="28e09-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="28e09-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="28e09-185">ReadItem</span></span>|
|[<span data-ttu-id="28e09-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="28e09-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="28e09-187">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="28e09-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28e09-188">例</span><span class="sxs-lookup"><span data-stu-id="28e09-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
