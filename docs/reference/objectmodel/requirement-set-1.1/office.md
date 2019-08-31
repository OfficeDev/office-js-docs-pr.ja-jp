---
title: Office 名前空間 - 要件セット 1.1
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: 70413bdfc01378bb5b1814fd938ab94a7e5101ba
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696394"
---
# <a name="office"></a><span data-ttu-id="d1197-102">Office</span><span class="sxs-lookup"><span data-stu-id="d1197-102">Office</span></span>

<span data-ttu-id="d1197-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d1197-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1197-105">要件</span><span class="sxs-lookup"><span data-stu-id="d1197-105">Requirements</span></span>

|<span data-ttu-id="d1197-106">要件</span><span class="sxs-lookup"><span data-stu-id="d1197-106">Requirement</span></span>| <span data-ttu-id="d1197-107">値</span><span class="sxs-lookup"><span data-stu-id="d1197-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1197-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1197-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1197-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d1197-109">1.0</span></span>|
|[<span data-ttu-id="d1197-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1197-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1197-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1197-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d1197-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="d1197-112">Members and methods</span></span>

| <span data-ttu-id="d1197-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1197-113">Member</span></span> | <span data-ttu-id="d1197-114">種類</span><span class="sxs-lookup"><span data-stu-id="d1197-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d1197-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="d1197-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="d1197-116">Member</span><span class="sxs-lookup"><span data-stu-id="d1197-116">Member</span></span> |
| [<span data-ttu-id="d1197-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="d1197-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="d1197-118">Member</span><span class="sxs-lookup"><span data-stu-id="d1197-118">Member</span></span> |
| [<span data-ttu-id="d1197-119">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="d1197-119">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="d1197-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1197-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="d1197-121">名前空間</span><span class="sxs-lookup"><span data-stu-id="d1197-121">Namespaces</span></span>

<span data-ttu-id="d1197-122">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="d1197-122">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="d1197-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.1):、、、、、、などのさまざま`ItemType`な`EntityType`列挙`AttachmentType` `RecipientType` `ResponseType`値が含まれ`ItemNotificationMessageType`ています。</span><span class="sxs-lookup"><span data-stu-id="d1197-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.1): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="d1197-124">Members</span><span class="sxs-lookup"><span data-stu-id="d1197-124">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="d1197-125">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="d1197-125">AsyncResultStatus: String</span></span>

<span data-ttu-id="d1197-126">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="d1197-126">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="d1197-127">型</span><span class="sxs-lookup"><span data-stu-id="d1197-127">Type</span></span>

*   <span data-ttu-id="d1197-128">String</span><span class="sxs-lookup"><span data-stu-id="d1197-128">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d1197-129">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="d1197-129">Properties:</span></span>

|<span data-ttu-id="d1197-130">名前</span><span class="sxs-lookup"><span data-stu-id="d1197-130">Name</span></span>| <span data-ttu-id="d1197-131">種類</span><span class="sxs-lookup"><span data-stu-id="d1197-131">Type</span></span>| <span data-ttu-id="d1197-132">説明</span><span class="sxs-lookup"><span data-stu-id="d1197-132">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="d1197-133">String</span><span class="sxs-lookup"><span data-stu-id="d1197-133">String</span></span>|<span data-ttu-id="d1197-134">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="d1197-134">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="d1197-135">String</span><span class="sxs-lookup"><span data-stu-id="d1197-135">String</span></span>|<span data-ttu-id="d1197-136">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="d1197-136">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d1197-137">要件</span><span class="sxs-lookup"><span data-stu-id="d1197-137">Requirements</span></span>

|<span data-ttu-id="d1197-138">要件</span><span class="sxs-lookup"><span data-stu-id="d1197-138">Requirement</span></span>| <span data-ttu-id="d1197-139">値</span><span class="sxs-lookup"><span data-stu-id="d1197-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1197-140">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1197-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1197-141">1.0</span><span class="sxs-lookup"><span data-stu-id="d1197-141">1.0</span></span>|
|[<span data-ttu-id="d1197-142">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1197-142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1197-143">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1197-143">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="d1197-144">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="d1197-144">CoercionType: String</span></span>

<span data-ttu-id="d1197-145">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="d1197-145">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d1197-146">型</span><span class="sxs-lookup"><span data-stu-id="d1197-146">Type</span></span>

*   <span data-ttu-id="d1197-147">String</span><span class="sxs-lookup"><span data-stu-id="d1197-147">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d1197-148">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="d1197-148">Properties:</span></span>

|<span data-ttu-id="d1197-149">名前</span><span class="sxs-lookup"><span data-stu-id="d1197-149">Name</span></span>| <span data-ttu-id="d1197-150">種類</span><span class="sxs-lookup"><span data-stu-id="d1197-150">Type</span></span>| <span data-ttu-id="d1197-151">説明</span><span class="sxs-lookup"><span data-stu-id="d1197-151">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="d1197-152">String</span><span class="sxs-lookup"><span data-stu-id="d1197-152">String</span></span>|<span data-ttu-id="d1197-153">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="d1197-153">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="d1197-154">String</span><span class="sxs-lookup"><span data-stu-id="d1197-154">String</span></span>|<span data-ttu-id="d1197-155">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="d1197-155">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d1197-156">要件</span><span class="sxs-lookup"><span data-stu-id="d1197-156">Requirements</span></span>

|<span data-ttu-id="d1197-157">要件</span><span class="sxs-lookup"><span data-stu-id="d1197-157">Requirement</span></span>| <span data-ttu-id="d1197-158">値</span><span class="sxs-lookup"><span data-stu-id="d1197-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1197-159">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1197-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1197-160">1.0</span><span class="sxs-lookup"><span data-stu-id="d1197-160">1.0</span></span>|
|[<span data-ttu-id="d1197-161">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1197-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1197-162">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1197-162">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="d1197-163">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="d1197-163">SourceProperty: String</span></span>

<span data-ttu-id="d1197-164">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="d1197-164">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d1197-165">型</span><span class="sxs-lookup"><span data-stu-id="d1197-165">Type</span></span>

*   <span data-ttu-id="d1197-166">String</span><span class="sxs-lookup"><span data-stu-id="d1197-166">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d1197-167">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="d1197-167">Properties:</span></span>

|<span data-ttu-id="d1197-168">名前</span><span class="sxs-lookup"><span data-stu-id="d1197-168">Name</span></span>| <span data-ttu-id="d1197-169">種類</span><span class="sxs-lookup"><span data-stu-id="d1197-169">Type</span></span>| <span data-ttu-id="d1197-170">説明</span><span class="sxs-lookup"><span data-stu-id="d1197-170">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="d1197-171">String</span><span class="sxs-lookup"><span data-stu-id="d1197-171">String</span></span>|<span data-ttu-id="d1197-172">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="d1197-172">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="d1197-173">String</span><span class="sxs-lookup"><span data-stu-id="d1197-173">String</span></span>|<span data-ttu-id="d1197-174">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="d1197-174">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d1197-175">要件</span><span class="sxs-lookup"><span data-stu-id="d1197-175">Requirements</span></span>

|<span data-ttu-id="d1197-176">要件</span><span class="sxs-lookup"><span data-stu-id="d1197-176">Requirement</span></span>| <span data-ttu-id="d1197-177">値</span><span class="sxs-lookup"><span data-stu-id="d1197-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1197-178">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1197-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1197-179">1.0</span><span class="sxs-lookup"><span data-stu-id="d1197-179">1.0</span></span>|
|[<span data-ttu-id="d1197-180">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1197-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1197-181">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1197-181">Compose or Read</span></span>|
