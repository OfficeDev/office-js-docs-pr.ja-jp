---
title: Office 名前空間-要件セット1.4
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: e6c4614af74a665805c400c407e4a7785efe9f96
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268650"
---
# <a name="office"></a><span data-ttu-id="3f43d-102">Office</span><span class="sxs-lookup"><span data-stu-id="3f43d-102">Office</span></span>

<span data-ttu-id="3f43d-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3f43d-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="3f43d-105">要件</span><span class="sxs-lookup"><span data-stu-id="3f43d-105">Requirements</span></span>

|<span data-ttu-id="3f43d-106">要件</span><span class="sxs-lookup"><span data-stu-id="3f43d-106">Requirement</span></span>| <span data-ttu-id="3f43d-107">値</span><span class="sxs-lookup"><span data-stu-id="3f43d-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f43d-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3f43d-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3f43d-109">1.0</span><span class="sxs-lookup"><span data-stu-id="3f43d-109">1.0</span></span>|
|[<span data-ttu-id="3f43d-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3f43d-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3f43d-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3f43d-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="3f43d-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="3f43d-112">Members and methods</span></span>

| <span data-ttu-id="3f43d-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="3f43d-113">Member</span></span> | <span data-ttu-id="3f43d-114">種類</span><span class="sxs-lookup"><span data-stu-id="3f43d-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="3f43d-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="3f43d-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="3f43d-116">Member</span><span class="sxs-lookup"><span data-stu-id="3f43d-116">Member</span></span> |
| [<span data-ttu-id="3f43d-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="3f43d-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="3f43d-118">Member</span><span class="sxs-lookup"><span data-stu-id="3f43d-118">Member</span></span> |
| [<span data-ttu-id="3f43d-119">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="3f43d-119">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="3f43d-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="3f43d-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="3f43d-121">名前空間</span><span class="sxs-lookup"><span data-stu-id="3f43d-121">Namespaces</span></span>

<span data-ttu-id="3f43d-122">[context](Office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="3f43d-122">[context](Office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="3f43d-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.4): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="3f43d-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.4): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="3f43d-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="3f43d-124">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="3f43d-125">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="3f43d-125">AsyncResultStatus: String</span></span>

<span data-ttu-id="3f43d-126">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="3f43d-126">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="3f43d-127">型</span><span class="sxs-lookup"><span data-stu-id="3f43d-127">Type</span></span>

*   <span data-ttu-id="3f43d-128">String</span><span class="sxs-lookup"><span data-stu-id="3f43d-128">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3f43d-129">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="3f43d-129">Properties:</span></span>

|<span data-ttu-id="3f43d-130">名前</span><span class="sxs-lookup"><span data-stu-id="3f43d-130">Name</span></span>| <span data-ttu-id="3f43d-131">種類</span><span class="sxs-lookup"><span data-stu-id="3f43d-131">Type</span></span>| <span data-ttu-id="3f43d-132">説明</span><span class="sxs-lookup"><span data-stu-id="3f43d-132">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="3f43d-133">String</span><span class="sxs-lookup"><span data-stu-id="3f43d-133">String</span></span>|<span data-ttu-id="3f43d-134">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="3f43d-134">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="3f43d-135">String</span><span class="sxs-lookup"><span data-stu-id="3f43d-135">String</span></span>|<span data-ttu-id="3f43d-136">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="3f43d-136">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3f43d-137">要件</span><span class="sxs-lookup"><span data-stu-id="3f43d-137">Requirements</span></span>

|<span data-ttu-id="3f43d-138">要件</span><span class="sxs-lookup"><span data-stu-id="3f43d-138">Requirement</span></span>| <span data-ttu-id="3f43d-139">値</span><span class="sxs-lookup"><span data-stu-id="3f43d-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f43d-140">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3f43d-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3f43d-141">1.0</span><span class="sxs-lookup"><span data-stu-id="3f43d-141">1.0</span></span>|
|[<span data-ttu-id="3f43d-142">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3f43d-142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3f43d-143">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3f43d-143">Compose or Read</span></span>|

#### <a name="coerciontype-string"></a><span data-ttu-id="3f43d-144">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="3f43d-144">CoercionType: String</span></span>

<span data-ttu-id="3f43d-145">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="3f43d-145">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3f43d-146">型</span><span class="sxs-lookup"><span data-stu-id="3f43d-146">Type</span></span>

*   <span data-ttu-id="3f43d-147">String</span><span class="sxs-lookup"><span data-stu-id="3f43d-147">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3f43d-148">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="3f43d-148">Properties:</span></span>

|<span data-ttu-id="3f43d-149">名前</span><span class="sxs-lookup"><span data-stu-id="3f43d-149">Name</span></span>| <span data-ttu-id="3f43d-150">種類</span><span class="sxs-lookup"><span data-stu-id="3f43d-150">Type</span></span>| <span data-ttu-id="3f43d-151">説明</span><span class="sxs-lookup"><span data-stu-id="3f43d-151">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="3f43d-152">String</span><span class="sxs-lookup"><span data-stu-id="3f43d-152">String</span></span>|<span data-ttu-id="3f43d-153">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="3f43d-153">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="3f43d-154">String</span><span class="sxs-lookup"><span data-stu-id="3f43d-154">String</span></span>|<span data-ttu-id="3f43d-155">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="3f43d-155">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3f43d-156">要件</span><span class="sxs-lookup"><span data-stu-id="3f43d-156">Requirements</span></span>

|<span data-ttu-id="3f43d-157">要件</span><span class="sxs-lookup"><span data-stu-id="3f43d-157">Requirement</span></span>| <span data-ttu-id="3f43d-158">値</span><span class="sxs-lookup"><span data-stu-id="3f43d-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f43d-159">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3f43d-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3f43d-160">1.0</span><span class="sxs-lookup"><span data-stu-id="3f43d-160">1.0</span></span>|
|[<span data-ttu-id="3f43d-161">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3f43d-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3f43d-162">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3f43d-162">Compose or Read</span></span>|

#### <a name="sourceproperty-string"></a><span data-ttu-id="3f43d-163">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="3f43d-163">SourceProperty: String</span></span>

<span data-ttu-id="3f43d-164">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="3f43d-164">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3f43d-165">型</span><span class="sxs-lookup"><span data-stu-id="3f43d-165">Type</span></span>

*   <span data-ttu-id="3f43d-166">String</span><span class="sxs-lookup"><span data-stu-id="3f43d-166">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3f43d-167">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="3f43d-167">Properties:</span></span>

|<span data-ttu-id="3f43d-168">名前</span><span class="sxs-lookup"><span data-stu-id="3f43d-168">Name</span></span>| <span data-ttu-id="3f43d-169">種類</span><span class="sxs-lookup"><span data-stu-id="3f43d-169">Type</span></span>| <span data-ttu-id="3f43d-170">説明</span><span class="sxs-lookup"><span data-stu-id="3f43d-170">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="3f43d-171">String</span><span class="sxs-lookup"><span data-stu-id="3f43d-171">String</span></span>|<span data-ttu-id="3f43d-172">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="3f43d-172">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="3f43d-173">String</span><span class="sxs-lookup"><span data-stu-id="3f43d-173">String</span></span>|<span data-ttu-id="3f43d-174">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="3f43d-174">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3f43d-175">要件</span><span class="sxs-lookup"><span data-stu-id="3f43d-175">Requirements</span></span>

|<span data-ttu-id="3f43d-176">要件</span><span class="sxs-lookup"><span data-stu-id="3f43d-176">Requirement</span></span>| <span data-ttu-id="3f43d-177">値</span><span class="sxs-lookup"><span data-stu-id="3f43d-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f43d-178">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3f43d-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3f43d-179">1.0</span><span class="sxs-lookup"><span data-stu-id="3f43d-179">1.0</span></span>|
|[<span data-ttu-id="3f43d-180">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3f43d-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3f43d-181">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3f43d-181">Compose or Read</span></span>|
