---
title: Office.context.mailbox.userProfile - プレビュー要件セット
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 061ee8367005f4af0795c4d9e1236d0b2443521a
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432817"
---
# <a name="userprofile"></a><span data-ttu-id="9595c-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="9595c-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="9595c-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="9595c-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="9595c-104">要件</span><span class="sxs-lookup"><span data-stu-id="9595c-104">Requirements</span></span>

|<span data-ttu-id="9595c-105">要件</span><span class="sxs-lookup"><span data-stu-id="9595c-105">Requirement</span></span>| <span data-ttu-id="9595c-106">値</span><span class="sxs-lookup"><span data-stu-id="9595c-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="9595c-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9595c-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9595c-108">1.0</span><span class="sxs-lookup"><span data-stu-id="9595c-108">1.0</span></span>|
|[<span data-ttu-id="9595c-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9595c-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9595c-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9595c-110">ReadItem</span></span>|
|[<span data-ttu-id="9595c-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9595c-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9595c-112">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="9595c-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="9595c-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="9595c-113">Members and methods</span></span>

| <span data-ttu-id="9595c-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="9595c-114">Member</span></span> | <span data-ttu-id="9595c-115">種類</span><span class="sxs-lookup"><span data-stu-id="9595c-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="9595c-116">accountType</span><span class="sxs-lookup"><span data-stu-id="9595c-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="9595c-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="9595c-117">Member</span></span> |
| [<span data-ttu-id="9595c-118">displayName</span><span class="sxs-lookup"><span data-stu-id="9595c-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="9595c-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="9595c-119">Member</span></span> |
| [<span data-ttu-id="9595c-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="9595c-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="9595c-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="9595c-121">Member</span></span> |
| [<span data-ttu-id="9595c-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="9595c-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="9595c-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="9595c-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="9595c-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="9595c-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="9595c-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="9595c-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="9595c-126">現在、このメンバーは Outlook 2016 for Mac 以降 (ビルド 16.9.1212 以降) でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="9595c-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="9595c-127">メールボックスに関連付けられているユーザーのアカウントの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="9595c-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="9595c-128">次の表に使用可能な値を示します。</span><span class="sxs-lookup"><span data-stu-id="9595c-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="9595c-129">値</span><span class="sxs-lookup"><span data-stu-id="9595c-129">Value</span></span> | <span data-ttu-id="9595c-130">説明</span><span class="sxs-lookup"><span data-stu-id="9595c-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="9595c-131">メールボックスは、オンプレミスの Exchange サーバーにあります。</span><span class="sxs-lookup"><span data-stu-id="9595c-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="9595c-132">メールボックスは、Gmail アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="9595c-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="9595c-133">メールボックスは、Office 365 の職場または学校のアカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="9595c-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="9595c-134">メールボックスは、個人の Outlook.com アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="9595c-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="9595c-135">種類:</span><span class="sxs-lookup"><span data-stu-id="9595c-135">Type:</span></span>

*   <span data-ttu-id="9595c-136">String</span><span class="sxs-lookup"><span data-stu-id="9595c-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9595c-137">要件</span><span class="sxs-lookup"><span data-stu-id="9595c-137">Requirements</span></span>

|<span data-ttu-id="9595c-138">要件</span><span class="sxs-lookup"><span data-stu-id="9595c-138">Requirement</span></span>| <span data-ttu-id="9595c-139">値</span><span class="sxs-lookup"><span data-stu-id="9595c-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="9595c-140">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9595c-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9595c-141">1.6</span><span class="sxs-lookup"><span data-stu-id="9595c-141">1.6</span></span> |
|[<span data-ttu-id="9595c-142">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9595c-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9595c-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9595c-143">ReadItem</span></span>|
|[<span data-ttu-id="9595c-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9595c-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9595c-145">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="9595c-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9595c-146">例</span><span class="sxs-lookup"><span data-stu-id="9595c-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="9595c-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="9595c-147">displayName :String</span></span>

<span data-ttu-id="9595c-148">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="9595c-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="9595c-149">型:</span><span class="sxs-lookup"><span data-stu-id="9595c-149">Type:</span></span>

*   <span data-ttu-id="9595c-150">String</span><span class="sxs-lookup"><span data-stu-id="9595c-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9595c-151">要件</span><span class="sxs-lookup"><span data-stu-id="9595c-151">Requirements</span></span>

|<span data-ttu-id="9595c-152">要件</span><span class="sxs-lookup"><span data-stu-id="9595c-152">Requirement</span></span>| <span data-ttu-id="9595c-153">値</span><span class="sxs-lookup"><span data-stu-id="9595c-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="9595c-154">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9595c-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9595c-155">1.0</span><span class="sxs-lookup"><span data-stu-id="9595c-155">1.0</span></span>|
|[<span data-ttu-id="9595c-156">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9595c-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9595c-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9595c-157">ReadItem</span></span>|
|[<span data-ttu-id="9595c-158">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9595c-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9595c-159">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="9595c-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9595c-160">例</span><span class="sxs-lookup"><span data-stu-id="9595c-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="9595c-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="9595c-161">emailAddress :String</span></span>

<span data-ttu-id="9595c-162">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="9595c-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="9595c-163">型:</span><span class="sxs-lookup"><span data-stu-id="9595c-163">Type:</span></span>

*   <span data-ttu-id="9595c-164">String</span><span class="sxs-lookup"><span data-stu-id="9595c-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9595c-165">要件</span><span class="sxs-lookup"><span data-stu-id="9595c-165">Requirements</span></span>

|<span data-ttu-id="9595c-166">要件</span><span class="sxs-lookup"><span data-stu-id="9595c-166">Requirement</span></span>| <span data-ttu-id="9595c-167">値</span><span class="sxs-lookup"><span data-stu-id="9595c-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="9595c-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9595c-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9595c-169">1.0</span><span class="sxs-lookup"><span data-stu-id="9595c-169">1.0</span></span>|
|[<span data-ttu-id="9595c-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9595c-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9595c-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9595c-171">ReadItem</span></span>|
|[<span data-ttu-id="9595c-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9595c-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9595c-173">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="9595c-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9595c-174">例</span><span class="sxs-lookup"><span data-stu-id="9595c-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="9595c-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="9595c-175">timeZone :String</span></span>

<span data-ttu-id="9595c-176">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="9595c-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="9595c-177">型:</span><span class="sxs-lookup"><span data-stu-id="9595c-177">Type:</span></span>

*   <span data-ttu-id="9595c-178">String</span><span class="sxs-lookup"><span data-stu-id="9595c-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9595c-179">要件</span><span class="sxs-lookup"><span data-stu-id="9595c-179">Requirements</span></span>

|<span data-ttu-id="9595c-180">要件</span><span class="sxs-lookup"><span data-stu-id="9595c-180">Requirement</span></span>| <span data-ttu-id="9595c-181">値</span><span class="sxs-lookup"><span data-stu-id="9595c-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="9595c-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9595c-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9595c-183">1.0</span><span class="sxs-lookup"><span data-stu-id="9595c-183">1.0</span></span>|
|[<span data-ttu-id="9595c-184">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="9595c-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9595c-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9595c-185">ReadItem</span></span>|
|[<span data-ttu-id="9595c-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9595c-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9595c-187">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="9595c-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9595c-188">例</span><span class="sxs-lookup"><span data-stu-id="9595c-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```