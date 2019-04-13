---
title: Office.-mailbox-要件セット1.7
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 8cfee874bbb5183d62cc3a9ce8b042a76617ec72
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838523"
---
# <a name="userprofile"></a><span data-ttu-id="200fb-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="200fb-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="200fb-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="200fb-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="200fb-104">要件</span><span class="sxs-lookup"><span data-stu-id="200fb-104">Requirements</span></span>

|<span data-ttu-id="200fb-105">要件</span><span class="sxs-lookup"><span data-stu-id="200fb-105">Requirement</span></span>| <span data-ttu-id="200fb-106">値</span><span class="sxs-lookup"><span data-stu-id="200fb-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="200fb-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="200fb-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="200fb-108">1.0</span><span class="sxs-lookup"><span data-stu-id="200fb-108">1.0</span></span>|
|[<span data-ttu-id="200fb-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="200fb-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="200fb-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="200fb-110">ReadItem</span></span>|
|[<span data-ttu-id="200fb-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="200fb-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="200fb-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="200fb-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="200fb-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="200fb-113">Members and methods</span></span>

| <span data-ttu-id="200fb-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="200fb-114">Member</span></span> | <span data-ttu-id="200fb-115">種類</span><span class="sxs-lookup"><span data-stu-id="200fb-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="200fb-116">accountType</span><span class="sxs-lookup"><span data-stu-id="200fb-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="200fb-117">Member</span><span class="sxs-lookup"><span data-stu-id="200fb-117">Member</span></span> |
| [<span data-ttu-id="200fb-118">displayName</span><span class="sxs-lookup"><span data-stu-id="200fb-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="200fb-119">Member</span><span class="sxs-lookup"><span data-stu-id="200fb-119">Member</span></span> |
| [<span data-ttu-id="200fb-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="200fb-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="200fb-121">Member</span><span class="sxs-lookup"><span data-stu-id="200fb-121">Member</span></span> |
| [<span data-ttu-id="200fb-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="200fb-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="200fb-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="200fb-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="200fb-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="200fb-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="200fb-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="200fb-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="200fb-126">このメンバーは、現在、Outlook 2016 for Mac (ビルド16.9.1212 以降) でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="200fb-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="200fb-127">メールボックスに関連付けられているユーザーのアカウントの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="200fb-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="200fb-128">次の表に使用可能な値を示します。</span><span class="sxs-lookup"><span data-stu-id="200fb-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="200fb-129">値</span><span class="sxs-lookup"><span data-stu-id="200fb-129">Value</span></span> | <span data-ttu-id="200fb-130">説明</span><span class="sxs-lookup"><span data-stu-id="200fb-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="200fb-131">メールボックスは、オンプレミスの Exchange サーバーにあります。</span><span class="sxs-lookup"><span data-stu-id="200fb-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="200fb-132">メールボックスは、Gmail アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="200fb-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="200fb-133">メールボックスは、Office 365 の職場または学校のアカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="200fb-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="200fb-134">メールボックスは、個人の Outlook.com アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="200fb-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="200fb-135">型</span><span class="sxs-lookup"><span data-stu-id="200fb-135">Type</span></span>

*   <span data-ttu-id="200fb-136">String</span><span class="sxs-lookup"><span data-stu-id="200fb-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="200fb-137">要件</span><span class="sxs-lookup"><span data-stu-id="200fb-137">Requirements</span></span>

|<span data-ttu-id="200fb-138">要件</span><span class="sxs-lookup"><span data-stu-id="200fb-138">Requirement</span></span>| <span data-ttu-id="200fb-139">値</span><span class="sxs-lookup"><span data-stu-id="200fb-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="200fb-140">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="200fb-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="200fb-141">1.6</span><span class="sxs-lookup"><span data-stu-id="200fb-141">1.6</span></span> |
|[<span data-ttu-id="200fb-142">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="200fb-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="200fb-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="200fb-143">ReadItem</span></span>|
|[<span data-ttu-id="200fb-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="200fb-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="200fb-145">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="200fb-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="200fb-146">例</span><span class="sxs-lookup"><span data-stu-id="200fb-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

---
---

####  <a name="displayname-string"></a><span data-ttu-id="200fb-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="200fb-147">displayName :String</span></span>

<span data-ttu-id="200fb-148">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="200fb-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="200fb-149">型</span><span class="sxs-lookup"><span data-stu-id="200fb-149">Type</span></span>

*   <span data-ttu-id="200fb-150">String</span><span class="sxs-lookup"><span data-stu-id="200fb-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="200fb-151">要件</span><span class="sxs-lookup"><span data-stu-id="200fb-151">Requirements</span></span>

|<span data-ttu-id="200fb-152">要件</span><span class="sxs-lookup"><span data-stu-id="200fb-152">Requirement</span></span>| <span data-ttu-id="200fb-153">値</span><span class="sxs-lookup"><span data-stu-id="200fb-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="200fb-154">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="200fb-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="200fb-155">1.0</span><span class="sxs-lookup"><span data-stu-id="200fb-155">1.0</span></span>|
|[<span data-ttu-id="200fb-156">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="200fb-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="200fb-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="200fb-157">ReadItem</span></span>|
|[<span data-ttu-id="200fb-158">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="200fb-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="200fb-159">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="200fb-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="200fb-160">例</span><span class="sxs-lookup"><span data-stu-id="200fb-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

---
---

####  <a name="emailaddress-string"></a><span data-ttu-id="200fb-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="200fb-161">emailAddress :String</span></span>

<span data-ttu-id="200fb-162">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="200fb-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="200fb-163">型</span><span class="sxs-lookup"><span data-stu-id="200fb-163">Type</span></span>

*   <span data-ttu-id="200fb-164">String</span><span class="sxs-lookup"><span data-stu-id="200fb-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="200fb-165">要件</span><span class="sxs-lookup"><span data-stu-id="200fb-165">Requirements</span></span>

|<span data-ttu-id="200fb-166">要件</span><span class="sxs-lookup"><span data-stu-id="200fb-166">Requirement</span></span>| <span data-ttu-id="200fb-167">値</span><span class="sxs-lookup"><span data-stu-id="200fb-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="200fb-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="200fb-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="200fb-169">1.0</span><span class="sxs-lookup"><span data-stu-id="200fb-169">1.0</span></span>|
|[<span data-ttu-id="200fb-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="200fb-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="200fb-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="200fb-171">ReadItem</span></span>|
|[<span data-ttu-id="200fb-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="200fb-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="200fb-173">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="200fb-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="200fb-174">例</span><span class="sxs-lookup"><span data-stu-id="200fb-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

---
---

####  <a name="timezone-string"></a><span data-ttu-id="200fb-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="200fb-175">timeZone :String</span></span>

<span data-ttu-id="200fb-176">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="200fb-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="200fb-177">型</span><span class="sxs-lookup"><span data-stu-id="200fb-177">Type</span></span>

*   <span data-ttu-id="200fb-178">String</span><span class="sxs-lookup"><span data-stu-id="200fb-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="200fb-179">要件</span><span class="sxs-lookup"><span data-stu-id="200fb-179">Requirements</span></span>

|<span data-ttu-id="200fb-180">要件</span><span class="sxs-lookup"><span data-stu-id="200fb-180">Requirement</span></span>| <span data-ttu-id="200fb-181">値</span><span class="sxs-lookup"><span data-stu-id="200fb-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="200fb-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="200fb-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="200fb-183">1.0</span><span class="sxs-lookup"><span data-stu-id="200fb-183">1.0</span></span>|
|[<span data-ttu-id="200fb-184">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="200fb-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="200fb-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="200fb-185">ReadItem</span></span>|
|[<span data-ttu-id="200fb-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="200fb-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="200fb-187">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="200fb-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="200fb-188">例</span><span class="sxs-lookup"><span data-stu-id="200fb-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
