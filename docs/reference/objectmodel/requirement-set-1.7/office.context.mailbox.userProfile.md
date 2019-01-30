---
title: Office.context.mailbox.userProfile - 要件セット 1.7
description: ''
ms.date: 10/31/2018
localization_priority: Normal
ms.openlocfilehash: b07ff5bee3adc18cc1006bb574e373182b29f5fe
ms.sourcegitcommit: 2e4b97f0252ff3dd908a3aa7a9720f0cb50b855d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/30/2019
ms.locfileid: "29635903"
---
# <a name="userprofile"></a><span data-ttu-id="ec496-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="ec496-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="ec496-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="ec496-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec496-104">要件</span><span class="sxs-lookup"><span data-stu-id="ec496-104">Requirements</span></span>

|<span data-ttu-id="ec496-105">要件</span><span class="sxs-lookup"><span data-stu-id="ec496-105">Requirement</span></span>| <span data-ttu-id="ec496-106">値</span><span class="sxs-lookup"><span data-stu-id="ec496-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec496-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ec496-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec496-108">1.0</span><span class="sxs-lookup"><span data-stu-id="ec496-108">1.0</span></span>|
|[<span data-ttu-id="ec496-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ec496-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec496-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec496-110">ReadItem</span></span>|
|[<span data-ttu-id="ec496-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ec496-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec496-112">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="ec496-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ec496-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="ec496-113">Members and methods</span></span>

| <span data-ttu-id="ec496-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="ec496-114">Member</span></span> | <span data-ttu-id="ec496-115">種類</span><span class="sxs-lookup"><span data-stu-id="ec496-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ec496-116">accountType</span><span class="sxs-lookup"><span data-stu-id="ec496-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="ec496-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="ec496-117">Member</span></span> |
| [<span data-ttu-id="ec496-118">displayName</span><span class="sxs-lookup"><span data-stu-id="ec496-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="ec496-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="ec496-119">Member</span></span> |
| [<span data-ttu-id="ec496-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="ec496-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="ec496-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="ec496-121">Member</span></span> |
| [<span data-ttu-id="ec496-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="ec496-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="ec496-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="ec496-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="ec496-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="ec496-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="ec496-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="ec496-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="ec496-126">このメンバーは現在、for Mac Outlook 2016 でのみサポートされている (ビルド 16.9.1212 またはそれ以降)。</span><span class="sxs-lookup"><span data-stu-id="ec496-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="ec496-127">メールボックスに関連付けられているユーザーのアカウントの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="ec496-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="ec496-128">次の表に使用可能な値を示します。</span><span class="sxs-lookup"><span data-stu-id="ec496-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="ec496-129">値</span><span class="sxs-lookup"><span data-stu-id="ec496-129">Value</span></span> | <span data-ttu-id="ec496-130">説明</span><span class="sxs-lookup"><span data-stu-id="ec496-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="ec496-131">メールボックスは、オンプレミスの Exchange サーバーにあります。</span><span class="sxs-lookup"><span data-stu-id="ec496-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="ec496-132">メールボックスは、Gmail アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="ec496-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="ec496-133">メールボックスは、Office 365 の職場または学校のアカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="ec496-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="ec496-134">メールボックスは、個人の Outlook.com アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="ec496-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="ec496-135">種類:</span><span class="sxs-lookup"><span data-stu-id="ec496-135">Type:</span></span>

*   <span data-ttu-id="ec496-136">String</span><span class="sxs-lookup"><span data-stu-id="ec496-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec496-137">要件</span><span class="sxs-lookup"><span data-stu-id="ec496-137">Requirements</span></span>

|<span data-ttu-id="ec496-138">要件</span><span class="sxs-lookup"><span data-stu-id="ec496-138">Requirement</span></span>| <span data-ttu-id="ec496-139">値</span><span class="sxs-lookup"><span data-stu-id="ec496-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec496-140">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ec496-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec496-141">1.6</span><span class="sxs-lookup"><span data-stu-id="ec496-141">1.6</span></span> |
|[<span data-ttu-id="ec496-142">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ec496-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec496-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec496-143">ReadItem</span></span>|
|[<span data-ttu-id="ec496-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ec496-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec496-145">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="ec496-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec496-146">例</span><span class="sxs-lookup"><span data-stu-id="ec496-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="ec496-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="ec496-147">displayName :String</span></span>

<span data-ttu-id="ec496-148">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="ec496-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="ec496-149">型:</span><span class="sxs-lookup"><span data-stu-id="ec496-149">Type:</span></span>

*   <span data-ttu-id="ec496-150">String</span><span class="sxs-lookup"><span data-stu-id="ec496-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec496-151">要件</span><span class="sxs-lookup"><span data-stu-id="ec496-151">Requirements</span></span>

|<span data-ttu-id="ec496-152">要件</span><span class="sxs-lookup"><span data-stu-id="ec496-152">Requirement</span></span>| <span data-ttu-id="ec496-153">値</span><span class="sxs-lookup"><span data-stu-id="ec496-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec496-154">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ec496-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec496-155">1.0</span><span class="sxs-lookup"><span data-stu-id="ec496-155">1.0</span></span>|
|[<span data-ttu-id="ec496-156">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ec496-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec496-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec496-157">ReadItem</span></span>|
|[<span data-ttu-id="ec496-158">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ec496-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec496-159">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="ec496-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec496-160">例</span><span class="sxs-lookup"><span data-stu-id="ec496-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="ec496-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="ec496-161">emailAddress :String</span></span>

<span data-ttu-id="ec496-162">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="ec496-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="ec496-163">型:</span><span class="sxs-lookup"><span data-stu-id="ec496-163">Type:</span></span>

*   <span data-ttu-id="ec496-164">String</span><span class="sxs-lookup"><span data-stu-id="ec496-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec496-165">要件</span><span class="sxs-lookup"><span data-stu-id="ec496-165">Requirements</span></span>

|<span data-ttu-id="ec496-166">要件</span><span class="sxs-lookup"><span data-stu-id="ec496-166">Requirement</span></span>| <span data-ttu-id="ec496-167">値</span><span class="sxs-lookup"><span data-stu-id="ec496-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec496-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ec496-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec496-169">1.0</span><span class="sxs-lookup"><span data-stu-id="ec496-169">1.0</span></span>|
|[<span data-ttu-id="ec496-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ec496-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec496-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec496-171">ReadItem</span></span>|
|[<span data-ttu-id="ec496-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ec496-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec496-173">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="ec496-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec496-174">例</span><span class="sxs-lookup"><span data-stu-id="ec496-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="ec496-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="ec496-175">timeZone :String</span></span>

<span data-ttu-id="ec496-176">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="ec496-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="ec496-177">型:</span><span class="sxs-lookup"><span data-stu-id="ec496-177">Type:</span></span>

*   <span data-ttu-id="ec496-178">String</span><span class="sxs-lookup"><span data-stu-id="ec496-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec496-179">要件</span><span class="sxs-lookup"><span data-stu-id="ec496-179">Requirements</span></span>

|<span data-ttu-id="ec496-180">要件</span><span class="sxs-lookup"><span data-stu-id="ec496-180">Requirement</span></span>| <span data-ttu-id="ec496-181">値</span><span class="sxs-lookup"><span data-stu-id="ec496-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec496-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ec496-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec496-183">1.0</span><span class="sxs-lookup"><span data-stu-id="ec496-183">1.0</span></span>|
|[<span data-ttu-id="ec496-184">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ec496-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ec496-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ec496-185">ReadItem</span></span>|
|[<span data-ttu-id="ec496-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ec496-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ec496-187">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="ec496-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ec496-188">例</span><span class="sxs-lookup"><span data-stu-id="ec496-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
