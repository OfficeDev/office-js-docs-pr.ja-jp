---
title: Office.context.mailbox.userProfile - 要件セット 1.6
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 09457a41fe68ae03e035d3d3f4b80b139be348e0
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067875"
---
# <a name="userprofile"></a><span data-ttu-id="e285e-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="e285e-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="e285e-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="e285e-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="e285e-104">要件</span><span class="sxs-lookup"><span data-stu-id="e285e-104">Requirements</span></span>

|<span data-ttu-id="e285e-105">要件</span><span class="sxs-lookup"><span data-stu-id="e285e-105">Requirement</span></span>| <span data-ttu-id="e285e-106">値</span><span class="sxs-lookup"><span data-stu-id="e285e-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="e285e-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e285e-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e285e-108">1.0</span><span class="sxs-lookup"><span data-stu-id="e285e-108">1.0</span></span>|
|[<span data-ttu-id="e285e-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e285e-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e285e-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e285e-110">ReadItem</span></span>|
|[<span data-ttu-id="e285e-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e285e-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e285e-112">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e285e-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e285e-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="e285e-113">Members and methods</span></span>

| <span data-ttu-id="e285e-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="e285e-114">Member</span></span> | <span data-ttu-id="e285e-115">種類</span><span class="sxs-lookup"><span data-stu-id="e285e-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e285e-116">accountType</span><span class="sxs-lookup"><span data-stu-id="e285e-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="e285e-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="e285e-117">Member</span></span> |
| [<span data-ttu-id="e285e-118">displayName</span><span class="sxs-lookup"><span data-stu-id="e285e-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="e285e-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="e285e-119">Member</span></span> |
| [<span data-ttu-id="e285e-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="e285e-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="e285e-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="e285e-121">Member</span></span> |
| [<span data-ttu-id="e285e-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="e285e-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="e285e-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="e285e-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="e285e-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="e285e-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="e285e-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="e285e-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="e285e-126">現在、このメンバーは Outlook 2016 for Mac 以降 (ビルド 16.9.1212 以降) でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="e285e-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="e285e-127">メールボックスに関連付けられているユーザーのアカウントの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="e285e-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="e285e-128">次の表に使用可能な値を示します。</span><span class="sxs-lookup"><span data-stu-id="e285e-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="e285e-129">値</span><span class="sxs-lookup"><span data-stu-id="e285e-129">Value</span></span> | <span data-ttu-id="e285e-130">説明</span><span class="sxs-lookup"><span data-stu-id="e285e-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="e285e-131">メールボックスは、オンプレミスの Exchange サーバーにあります。</span><span class="sxs-lookup"><span data-stu-id="e285e-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="e285e-132">メールボックスは、Gmail アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="e285e-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="e285e-133">メールボックスは、Office 365 の職場または学校のアカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="e285e-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="e285e-134">メールボックスは、個人の Outlook.com アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="e285e-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="e285e-135">Type</span><span class="sxs-lookup"><span data-stu-id="e285e-135">Type</span></span>

*   <span data-ttu-id="e285e-136">String</span><span class="sxs-lookup"><span data-stu-id="e285e-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e285e-137">要件</span><span class="sxs-lookup"><span data-stu-id="e285e-137">Requirements</span></span>

|<span data-ttu-id="e285e-138">要件</span><span class="sxs-lookup"><span data-stu-id="e285e-138">Requirement</span></span>| <span data-ttu-id="e285e-139">値</span><span class="sxs-lookup"><span data-stu-id="e285e-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="e285e-140">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e285e-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e285e-141">1.6</span><span class="sxs-lookup"><span data-stu-id="e285e-141">1.6</span></span> |
|[<span data-ttu-id="e285e-142">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e285e-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e285e-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e285e-143">ReadItem</span></span>|
|[<span data-ttu-id="e285e-144">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e285e-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e285e-145">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e285e-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e285e-146">例</span><span class="sxs-lookup"><span data-stu-id="e285e-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="e285e-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="e285e-147">displayName :String</span></span>

<span data-ttu-id="e285e-148">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="e285e-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="e285e-149">Type</span><span class="sxs-lookup"><span data-stu-id="e285e-149">Type</span></span>

*   <span data-ttu-id="e285e-150">String</span><span class="sxs-lookup"><span data-stu-id="e285e-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e285e-151">要件</span><span class="sxs-lookup"><span data-stu-id="e285e-151">Requirements</span></span>

|<span data-ttu-id="e285e-152">要件</span><span class="sxs-lookup"><span data-stu-id="e285e-152">Requirement</span></span>| <span data-ttu-id="e285e-153">値</span><span class="sxs-lookup"><span data-stu-id="e285e-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="e285e-154">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e285e-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e285e-155">1.0</span><span class="sxs-lookup"><span data-stu-id="e285e-155">1.0</span></span>|
|[<span data-ttu-id="e285e-156">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e285e-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e285e-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e285e-157">ReadItem</span></span>|
|[<span data-ttu-id="e285e-158">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e285e-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e285e-159">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e285e-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e285e-160">例</span><span class="sxs-lookup"><span data-stu-id="e285e-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="e285e-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="e285e-161">emailAddress :String</span></span>

<span data-ttu-id="e285e-162">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="e285e-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="e285e-163">Type</span><span class="sxs-lookup"><span data-stu-id="e285e-163">Type</span></span>

*   <span data-ttu-id="e285e-164">String</span><span class="sxs-lookup"><span data-stu-id="e285e-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e285e-165">要件</span><span class="sxs-lookup"><span data-stu-id="e285e-165">Requirements</span></span>

|<span data-ttu-id="e285e-166">要件</span><span class="sxs-lookup"><span data-stu-id="e285e-166">Requirement</span></span>| <span data-ttu-id="e285e-167">値</span><span class="sxs-lookup"><span data-stu-id="e285e-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="e285e-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e285e-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e285e-169">1.0</span><span class="sxs-lookup"><span data-stu-id="e285e-169">1.0</span></span>|
|[<span data-ttu-id="e285e-170">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e285e-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e285e-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e285e-171">ReadItem</span></span>|
|[<span data-ttu-id="e285e-172">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e285e-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e285e-173">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e285e-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e285e-174">例</span><span class="sxs-lookup"><span data-stu-id="e285e-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="e285e-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="e285e-175">timeZone :String</span></span>

<span data-ttu-id="e285e-176">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="e285e-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="e285e-177">Type</span><span class="sxs-lookup"><span data-stu-id="e285e-177">Type</span></span>

*   <span data-ttu-id="e285e-178">String</span><span class="sxs-lookup"><span data-stu-id="e285e-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e285e-179">要件</span><span class="sxs-lookup"><span data-stu-id="e285e-179">Requirements</span></span>

|<span data-ttu-id="e285e-180">要件</span><span class="sxs-lookup"><span data-stu-id="e285e-180">Requirement</span></span>| <span data-ttu-id="e285e-181">値</span><span class="sxs-lookup"><span data-stu-id="e285e-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="e285e-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e285e-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e285e-183">1.0</span><span class="sxs-lookup"><span data-stu-id="e285e-183">1.0</span></span>|
|[<span data-ttu-id="e285e-184">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e285e-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e285e-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e285e-185">ReadItem</span></span>|
|[<span data-ttu-id="e285e-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e285e-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e285e-187">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e285e-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e285e-188">例</span><span class="sxs-lookup"><span data-stu-id="e285e-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
