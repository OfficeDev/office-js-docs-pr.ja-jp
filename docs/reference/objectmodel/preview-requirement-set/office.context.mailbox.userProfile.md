---
title: Office.context.mailbox.userProfile - プレビュー要件セット
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 4afc64f247155576ab3f0024d1929a29a0f7dc0c
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629259"
---
# <a name="userprofile"></a><span data-ttu-id="6d8b9-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="6d8b9-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="6d8b9-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="6d8b9-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d8b9-104">Requirements</span><span class="sxs-lookup"><span data-stu-id="6d8b9-104">Requirements</span></span>

|<span data-ttu-id="6d8b9-105">要件</span><span class="sxs-lookup"><span data-stu-id="6d8b9-105">Requirement</span></span>| <span data-ttu-id="6d8b9-106">値</span><span class="sxs-lookup"><span data-stu-id="6d8b9-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d8b9-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6d8b9-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d8b9-108">1.0</span><span class="sxs-lookup"><span data-stu-id="6d8b9-108">1.0</span></span>|
|[<span data-ttu-id="6d8b9-109">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6d8b9-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d8b9-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d8b9-110">ReadItem</span></span>|
|[<span data-ttu-id="6d8b9-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6d8b9-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d8b9-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6d8b9-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="6d8b9-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="6d8b9-113">Properties</span></span>

| <span data-ttu-id="6d8b9-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="6d8b9-114">Property</span></span> | <span data-ttu-id="6d8b9-115">最小値</span><span class="sxs-lookup"><span data-stu-id="6d8b9-115">Minimum</span></span><br><span data-ttu-id="6d8b9-116">アクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6d8b9-116">permission level</span></span> | <span data-ttu-id="6d8b9-117">モード</span><span class="sxs-lookup"><span data-stu-id="6d8b9-117">Modes</span></span> | <span data-ttu-id="6d8b9-118">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="6d8b9-118">Return type</span></span> | <span data-ttu-id="6d8b9-119">最小値</span><span class="sxs-lookup"><span data-stu-id="6d8b9-119">Minimum</span></span><br><span data-ttu-id="6d8b9-120">要件セット</span><span class="sxs-lookup"><span data-stu-id="6d8b9-120">requirement set</span></span> |
|---|---|---|---|---|
| [<span data-ttu-id="6d8b9-121">accountType</span><span class="sxs-lookup"><span data-stu-id="6d8b9-121">accountType</span></span>](#accounttype-string) | <span data-ttu-id="6d8b9-122">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d8b9-122">ReadItem</span></span> | <span data-ttu-id="6d8b9-123">作成</span><span class="sxs-lookup"><span data-stu-id="6d8b9-123">Compose</span></span><br><span data-ttu-id="6d8b9-124">読み取り</span><span class="sxs-lookup"><span data-stu-id="6d8b9-124">Read</span></span> | <span data-ttu-id="6d8b9-125">String</span><span class="sxs-lookup"><span data-stu-id="6d8b9-125">String</span></span> | <span data-ttu-id="6d8b9-126">1.6</span><span class="sxs-lookup"><span data-stu-id="6d8b9-126">1.6</span></span> |
| [<span data-ttu-id="6d8b9-127">displayName</span><span class="sxs-lookup"><span data-stu-id="6d8b9-127">displayName</span></span>](#displayname-string) | <span data-ttu-id="6d8b9-128">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d8b9-128">ReadItem</span></span> | <span data-ttu-id="6d8b9-129">作成</span><span class="sxs-lookup"><span data-stu-id="6d8b9-129">Compose</span></span><br><span data-ttu-id="6d8b9-130">読み取り</span><span class="sxs-lookup"><span data-stu-id="6d8b9-130">Read</span></span> | <span data-ttu-id="6d8b9-131">String</span><span class="sxs-lookup"><span data-stu-id="6d8b9-131">String</span></span> | <span data-ttu-id="6d8b9-132">1.0</span><span class="sxs-lookup"><span data-stu-id="6d8b9-132">1.0</span></span> |
| [<span data-ttu-id="6d8b9-133">emailAddress</span><span class="sxs-lookup"><span data-stu-id="6d8b9-133">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="6d8b9-134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d8b9-134">ReadItem</span></span> | <span data-ttu-id="6d8b9-135">作成</span><span class="sxs-lookup"><span data-stu-id="6d8b9-135">Compose</span></span><br><span data-ttu-id="6d8b9-136">読み取り</span><span class="sxs-lookup"><span data-stu-id="6d8b9-136">Read</span></span> | <span data-ttu-id="6d8b9-137">String</span><span class="sxs-lookup"><span data-stu-id="6d8b9-137">String</span></span> | <span data-ttu-id="6d8b9-138">1.0</span><span class="sxs-lookup"><span data-stu-id="6d8b9-138">1.0</span></span> |
| [<span data-ttu-id="6d8b9-139">timeZone</span><span class="sxs-lookup"><span data-stu-id="6d8b9-139">timeZone</span></span>](#timezone-string) | <span data-ttu-id="6d8b9-140">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d8b9-140">ReadItem</span></span> | <span data-ttu-id="6d8b9-141">作成</span><span class="sxs-lookup"><span data-stu-id="6d8b9-141">Compose</span></span><br><span data-ttu-id="6d8b9-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="6d8b9-142">Read</span></span> | <span data-ttu-id="6d8b9-143">String</span><span class="sxs-lookup"><span data-stu-id="6d8b9-143">String</span></span> | <span data-ttu-id="6d8b9-144">1.0</span><span class="sxs-lookup"><span data-stu-id="6d8b9-144">1.0</span></span> |

## <a name="property-details"></a><span data-ttu-id="6d8b9-145">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="6d8b9-145">Property details</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="6d8b9-146">accountType: String</span><span class="sxs-lookup"><span data-stu-id="6d8b9-146">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="6d8b9-147">このメンバーは、現在、Outlook 2016 以降の Mac (ビルド16.9.1212 以降) でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="6d8b9-147">This member is currently only supported in Outlook 2016 or later on Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="6d8b9-148">メールボックスに関連付けられているユーザーのアカウントの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="6d8b9-148">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="6d8b9-149">次の表に使用可能な値を示します。</span><span class="sxs-lookup"><span data-stu-id="6d8b9-149">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="6d8b9-150">値</span><span class="sxs-lookup"><span data-stu-id="6d8b9-150">Value</span></span> | <span data-ttu-id="6d8b9-151">説明</span><span class="sxs-lookup"><span data-stu-id="6d8b9-151">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="6d8b9-152">メールボックスは、オンプレミスの Exchange サーバーにあります。</span><span class="sxs-lookup"><span data-stu-id="6d8b9-152">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="6d8b9-153">メールボックスは、Gmail アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="6d8b9-153">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="6d8b9-154">メールボックスは、Office 365 の職場または学校のアカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="6d8b9-154">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="6d8b9-155">メールボックスは、個人の Outlook.com アカウントに関連付けられます。</span><span class="sxs-lookup"><span data-stu-id="6d8b9-155">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="6d8b9-156">型</span><span class="sxs-lookup"><span data-stu-id="6d8b9-156">Type</span></span>

*   <span data-ttu-id="6d8b9-157">String</span><span class="sxs-lookup"><span data-stu-id="6d8b9-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d8b9-158">要件</span><span class="sxs-lookup"><span data-stu-id="6d8b9-158">Requirements</span></span>

|<span data-ttu-id="6d8b9-159">要件</span><span class="sxs-lookup"><span data-stu-id="6d8b9-159">Requirement</span></span>| <span data-ttu-id="6d8b9-160">値</span><span class="sxs-lookup"><span data-stu-id="6d8b9-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d8b9-161">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6d8b9-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d8b9-162">1.6</span><span class="sxs-lookup"><span data-stu-id="6d8b9-162">1.6</span></span> |
|[<span data-ttu-id="6d8b9-163">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6d8b9-163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d8b9-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d8b9-164">ReadItem</span></span>|
|[<span data-ttu-id="6d8b9-165">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6d8b9-165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d8b9-166">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6d8b9-166">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d8b9-167">例</span><span class="sxs-lookup"><span data-stu-id="6d8b9-167">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

<br>

---
---

#### <a name="displayname-string"></a><span data-ttu-id="6d8b9-168">displayName: String</span><span class="sxs-lookup"><span data-stu-id="6d8b9-168">displayName: String</span></span>

<span data-ttu-id="6d8b9-169">ユーザーの表示名を取得します。</span><span class="sxs-lookup"><span data-stu-id="6d8b9-169">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="6d8b9-170">型</span><span class="sxs-lookup"><span data-stu-id="6d8b9-170">Type</span></span>

*   <span data-ttu-id="6d8b9-171">String</span><span class="sxs-lookup"><span data-stu-id="6d8b9-171">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d8b9-172">要件</span><span class="sxs-lookup"><span data-stu-id="6d8b9-172">Requirements</span></span>

|<span data-ttu-id="6d8b9-173">要件</span><span class="sxs-lookup"><span data-stu-id="6d8b9-173">Requirement</span></span>| <span data-ttu-id="6d8b9-174">値</span><span class="sxs-lookup"><span data-stu-id="6d8b9-174">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d8b9-175">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6d8b9-175">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d8b9-176">1.0</span><span class="sxs-lookup"><span data-stu-id="6d8b9-176">1.0</span></span>|
|[<span data-ttu-id="6d8b9-177">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6d8b9-177">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d8b9-178">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d8b9-178">ReadItem</span></span>|
|[<span data-ttu-id="6d8b9-179">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6d8b9-179">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d8b9-180">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6d8b9-180">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d8b9-181">例</span><span class="sxs-lookup"><span data-stu-id="6d8b9-181">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="6d8b9-182">emailAddress: String</span><span class="sxs-lookup"><span data-stu-id="6d8b9-182">emailAddress: String</span></span>

<span data-ttu-id="6d8b9-183">ユーザーの SMTP 電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="6d8b9-183">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="6d8b9-184">型</span><span class="sxs-lookup"><span data-stu-id="6d8b9-184">Type</span></span>

*   <span data-ttu-id="6d8b9-185">String</span><span class="sxs-lookup"><span data-stu-id="6d8b9-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d8b9-186">要件</span><span class="sxs-lookup"><span data-stu-id="6d8b9-186">Requirements</span></span>

|<span data-ttu-id="6d8b9-187">要件</span><span class="sxs-lookup"><span data-stu-id="6d8b9-187">Requirement</span></span>| <span data-ttu-id="6d8b9-188">値</span><span class="sxs-lookup"><span data-stu-id="6d8b9-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d8b9-189">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6d8b9-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d8b9-190">1.0</span><span class="sxs-lookup"><span data-stu-id="6d8b9-190">1.0</span></span>|
|[<span data-ttu-id="6d8b9-191">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6d8b9-191">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d8b9-192">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d8b9-192">ReadItem</span></span>|
|[<span data-ttu-id="6d8b9-193">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6d8b9-193">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d8b9-194">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6d8b9-194">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d8b9-195">例</span><span class="sxs-lookup"><span data-stu-id="6d8b9-195">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="6d8b9-196">timeZone: String</span><span class="sxs-lookup"><span data-stu-id="6d8b9-196">timeZone: String</span></span>

<span data-ttu-id="6d8b9-197">ユーザーの既定のタイム ゾーンを取得します。</span><span class="sxs-lookup"><span data-stu-id="6d8b9-197">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="6d8b9-198">型</span><span class="sxs-lookup"><span data-stu-id="6d8b9-198">Type</span></span>

*   <span data-ttu-id="6d8b9-199">String</span><span class="sxs-lookup"><span data-stu-id="6d8b9-199">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d8b9-200">要件</span><span class="sxs-lookup"><span data-stu-id="6d8b9-200">Requirements</span></span>

|<span data-ttu-id="6d8b9-201">要件</span><span class="sxs-lookup"><span data-stu-id="6d8b9-201">Requirement</span></span>| <span data-ttu-id="6d8b9-202">値</span><span class="sxs-lookup"><span data-stu-id="6d8b9-202">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d8b9-203">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6d8b9-203">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d8b9-204">1.0</span><span class="sxs-lookup"><span data-stu-id="6d8b9-204">1.0</span></span>|
|[<span data-ttu-id="6d8b9-205">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6d8b9-205">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d8b9-206">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d8b9-206">ReadItem</span></span>|
|[<span data-ttu-id="6d8b9-207">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6d8b9-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d8b9-208">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6d8b9-208">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d8b9-209">例</span><span class="sxs-lookup"><span data-stu-id="6d8b9-209">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
