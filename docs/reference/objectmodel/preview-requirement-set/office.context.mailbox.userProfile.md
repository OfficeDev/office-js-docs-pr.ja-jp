---
title: Office.context.mailbox.userProfile - プレビュー要件セット
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 214434c988c01ecb1aef93f4067cd95bfe768ae9
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068176"
---
# <a name="userprofile"></a><span data-ttu-id="3a7b2-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="3a7b2-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="3a7b2-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="3a7b2-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a7b2-104">要件</span><span class="sxs-lookup"><span data-stu-id="3a7b2-104">Requirements</span></span>

|<span data-ttu-id="3a7b2-105">要件</span><span class="sxs-lookup"><span data-stu-id="3a7b2-105">Requirement</span></span>| <span data-ttu-id="3a7b2-106">値</span><span class="sxs-lookup"><span data-stu-id="3a7b2-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a7b2-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a7b2-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a7b2-108">1.0</span><span class="sxs-lookup"><span data-stu-id="3a7b2-108">1.0</span></span>|
|[<span data-ttu-id="3a7b2-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="3a7b2-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a7b2-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a7b2-110">ReadItem</span></span>|
|[<span data-ttu-id="3a7b2-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a7b2-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a7b2-112">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3a7b2-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="3a7b2-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="3a7b2-113">Members and methods</span></span>

| <span data-ttu-id="3a7b2-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a7b2-114">Member</span></span> | <span data-ttu-id="3a7b2-115">種類</span><span class="sxs-lookup"><span data-stu-id="3a7b2-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="3a7b2-116">accountType</span><span class="sxs-lookup"><span data-stu-id="3a7b2-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="3a7b2-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a7b2-117">Member</span></span> |
| [<span data-ttu-id="3a7b2-118">displayName</span><span class="sxs-lookup"><span data-stu-id="3a7b2-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="3a7b2-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a7b2-119">Member</span></span> |
| [<span data-ttu-id="3a7b2-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="3a7b2-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="3a7b2-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a7b2-121">Member</span></span> |
| [<span data-ttu-id="3a7b2-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="3a7b2-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="3a7b2-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a7b2-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="3a7b2-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a7b2-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="3a7b2-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="3a7b2-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="3a7b2-126">現在、このメンバーは Outlook 2016 for Mac 以降 (ビルド 16.9.1212 以降) でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="3a7b2-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="3a7b2-127">メールボックスに関連付けられているユーザーのアカウントの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="3a7b2-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="3a7b2-128">次の表に使用可能な値を示します。</span><span class="sxs-lookup"><span data-stu-id="3a7b2-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="3a7b2-129">値</span><span class="sxs-lookup"><span data-stu-id="3a7b2-129">Value</span></span> | <span data-ttu-id="3a7b2-130">説明</span><span class="sxs-lookup"><span data-stu-id="3a7b2-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="3a7b2-131">メールボックスは、オンプレミスの Exchange サーバーにあります。</span><span class="sxs-lookup"><span data-stu-id="3a7b2-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="3a7b2-132">メールボックスは、Gmail アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="3a7b2-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="3a7b2-133">メールボックスは、Office 365 の職場または学校のアカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="3a7b2-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="3a7b2-134">メールボックスは、個人の Outlook.com アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="3a7b2-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="3a7b2-135">Type</span><span class="sxs-lookup"><span data-stu-id="3a7b2-135">Type</span></span>

*   <span data-ttu-id="3a7b2-136">String</span><span class="sxs-lookup"><span data-stu-id="3a7b2-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a7b2-137">要件</span><span class="sxs-lookup"><span data-stu-id="3a7b2-137">Requirements</span></span>

|<span data-ttu-id="3a7b2-138">要件</span><span class="sxs-lookup"><span data-stu-id="3a7b2-138">Requirement</span></span>| <span data-ttu-id="3a7b2-139">値</span><span class="sxs-lookup"><span data-stu-id="3a7b2-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a7b2-140">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a7b2-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a7b2-141">1.6</span><span class="sxs-lookup"><span data-stu-id="3a7b2-141">1.6</span></span> |
|[<span data-ttu-id="3a7b2-142">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="3a7b2-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a7b2-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a7b2-143">ReadItem</span></span>|
|[<span data-ttu-id="3a7b2-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a7b2-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a7b2-145">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3a7b2-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a7b2-146">例</span><span class="sxs-lookup"><span data-stu-id="3a7b2-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="3a7b2-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="3a7b2-147">displayName :String</span></span>

<span data-ttu-id="3a7b2-148">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="3a7b2-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="3a7b2-149">Type</span><span class="sxs-lookup"><span data-stu-id="3a7b2-149">Type</span></span>

*   <span data-ttu-id="3a7b2-150">String</span><span class="sxs-lookup"><span data-stu-id="3a7b2-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a7b2-151">要件</span><span class="sxs-lookup"><span data-stu-id="3a7b2-151">Requirements</span></span>

|<span data-ttu-id="3a7b2-152">要件</span><span class="sxs-lookup"><span data-stu-id="3a7b2-152">Requirement</span></span>| <span data-ttu-id="3a7b2-153">値</span><span class="sxs-lookup"><span data-stu-id="3a7b2-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a7b2-154">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a7b2-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a7b2-155">1.0</span><span class="sxs-lookup"><span data-stu-id="3a7b2-155">1.0</span></span>|
|[<span data-ttu-id="3a7b2-156">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="3a7b2-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a7b2-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a7b2-157">ReadItem</span></span>|
|[<span data-ttu-id="3a7b2-158">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a7b2-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a7b2-159">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3a7b2-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a7b2-160">例</span><span class="sxs-lookup"><span data-stu-id="3a7b2-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="3a7b2-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="3a7b2-161">emailAddress :String</span></span>

<span data-ttu-id="3a7b2-162">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="3a7b2-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="3a7b2-163">Type</span><span class="sxs-lookup"><span data-stu-id="3a7b2-163">Type</span></span>

*   <span data-ttu-id="3a7b2-164">String</span><span class="sxs-lookup"><span data-stu-id="3a7b2-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a7b2-165">要件</span><span class="sxs-lookup"><span data-stu-id="3a7b2-165">Requirements</span></span>

|<span data-ttu-id="3a7b2-166">要件</span><span class="sxs-lookup"><span data-stu-id="3a7b2-166">Requirement</span></span>| <span data-ttu-id="3a7b2-167">値</span><span class="sxs-lookup"><span data-stu-id="3a7b2-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a7b2-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a7b2-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a7b2-169">1.0</span><span class="sxs-lookup"><span data-stu-id="3a7b2-169">1.0</span></span>|
|[<span data-ttu-id="3a7b2-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="3a7b2-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a7b2-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a7b2-171">ReadItem</span></span>|
|[<span data-ttu-id="3a7b2-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a7b2-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a7b2-173">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3a7b2-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a7b2-174">例</span><span class="sxs-lookup"><span data-stu-id="3a7b2-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="3a7b2-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="3a7b2-175">timeZone :String</span></span>

<span data-ttu-id="3a7b2-176">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="3a7b2-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="3a7b2-177">Type</span><span class="sxs-lookup"><span data-stu-id="3a7b2-177">Type</span></span>

*   <span data-ttu-id="3a7b2-178">String</span><span class="sxs-lookup"><span data-stu-id="3a7b2-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a7b2-179">要件</span><span class="sxs-lookup"><span data-stu-id="3a7b2-179">Requirements</span></span>

|<span data-ttu-id="3a7b2-180">要件</span><span class="sxs-lookup"><span data-stu-id="3a7b2-180">Requirement</span></span>| <span data-ttu-id="3a7b2-181">値</span><span class="sxs-lookup"><span data-stu-id="3a7b2-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a7b2-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a7b2-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a7b2-183">1.0</span><span class="sxs-lookup"><span data-stu-id="3a7b2-183">1.0</span></span>|
|[<span data-ttu-id="3a7b2-184">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="3a7b2-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a7b2-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a7b2-185">ReadItem</span></span>|
|[<span data-ttu-id="3a7b2-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a7b2-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a7b2-187">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3a7b2-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a7b2-188">例</span><span class="sxs-lookup"><span data-stu-id="3a7b2-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
