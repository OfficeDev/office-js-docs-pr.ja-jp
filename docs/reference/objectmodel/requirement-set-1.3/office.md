---
title: Office 名前空間-要件セット1.3
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 0b22574693fb129be6a08a89b58beceb746fa283
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268398"
---
# <a name="office"></a><span data-ttu-id="b9660-102">Office</span><span class="sxs-lookup"><span data-stu-id="b9660-102">Office</span></span>

<span data-ttu-id="b9660-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b9660-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9660-105">要件</span><span class="sxs-lookup"><span data-stu-id="b9660-105">Requirements</span></span>

|<span data-ttu-id="b9660-106">要件</span><span class="sxs-lookup"><span data-stu-id="b9660-106">Requirement</span></span>| <span data-ttu-id="b9660-107">値</span><span class="sxs-lookup"><span data-stu-id="b9660-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9660-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b9660-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9660-109">1.0</span><span class="sxs-lookup"><span data-stu-id="b9660-109">1.0</span></span>|
|[<span data-ttu-id="b9660-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b9660-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9660-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b9660-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b9660-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="b9660-112">Members and methods</span></span>

| <span data-ttu-id="b9660-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="b9660-113">Member</span></span> | <span data-ttu-id="b9660-114">種類</span><span class="sxs-lookup"><span data-stu-id="b9660-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b9660-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="b9660-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="b9660-116">Member</span><span class="sxs-lookup"><span data-stu-id="b9660-116">Member</span></span> |
| [<span data-ttu-id="b9660-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="b9660-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="b9660-118">Member</span><span class="sxs-lookup"><span data-stu-id="b9660-118">Member</span></span> |
| [<span data-ttu-id="b9660-119">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="b9660-119">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="b9660-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="b9660-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="b9660-121">名前空間</span><span class="sxs-lookup"><span data-stu-id="b9660-121">Namespaces</span></span>

<span data-ttu-id="b9660-122">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="b9660-122">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="b9660-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.3): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="b9660-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.3): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="b9660-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="b9660-124">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="b9660-125">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="b9660-125">AsyncResultStatus: String</span></span>

<span data-ttu-id="b9660-126">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="b9660-126">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="b9660-127">型</span><span class="sxs-lookup"><span data-stu-id="b9660-127">Type</span></span>

*   <span data-ttu-id="b9660-128">String</span><span class="sxs-lookup"><span data-stu-id="b9660-128">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b9660-129">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="b9660-129">Properties:</span></span>

|<span data-ttu-id="b9660-130">名前</span><span class="sxs-lookup"><span data-stu-id="b9660-130">Name</span></span>| <span data-ttu-id="b9660-131">種類</span><span class="sxs-lookup"><span data-stu-id="b9660-131">Type</span></span>| <span data-ttu-id="b9660-132">説明</span><span class="sxs-lookup"><span data-stu-id="b9660-132">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="b9660-133">String</span><span class="sxs-lookup"><span data-stu-id="b9660-133">String</span></span>|<span data-ttu-id="b9660-134">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="b9660-134">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="b9660-135">String</span><span class="sxs-lookup"><span data-stu-id="b9660-135">String</span></span>|<span data-ttu-id="b9660-136">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="b9660-136">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b9660-137">要件</span><span class="sxs-lookup"><span data-stu-id="b9660-137">Requirements</span></span>

|<span data-ttu-id="b9660-138">要件</span><span class="sxs-lookup"><span data-stu-id="b9660-138">Requirement</span></span>| <span data-ttu-id="b9660-139">値</span><span class="sxs-lookup"><span data-stu-id="b9660-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9660-140">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b9660-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9660-141">1.0</span><span class="sxs-lookup"><span data-stu-id="b9660-141">1.0</span></span>|
|[<span data-ttu-id="b9660-142">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b9660-142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9660-143">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b9660-143">Compose or Read</span></span>|

#### <a name="coerciontype-string"></a><span data-ttu-id="b9660-144">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="b9660-144">CoercionType: String</span></span>

<span data-ttu-id="b9660-145">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="b9660-145">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b9660-146">型</span><span class="sxs-lookup"><span data-stu-id="b9660-146">Type</span></span>

*   <span data-ttu-id="b9660-147">String</span><span class="sxs-lookup"><span data-stu-id="b9660-147">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b9660-148">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="b9660-148">Properties:</span></span>

|<span data-ttu-id="b9660-149">名前</span><span class="sxs-lookup"><span data-stu-id="b9660-149">Name</span></span>| <span data-ttu-id="b9660-150">種類</span><span class="sxs-lookup"><span data-stu-id="b9660-150">Type</span></span>| <span data-ttu-id="b9660-151">説明</span><span class="sxs-lookup"><span data-stu-id="b9660-151">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="b9660-152">String</span><span class="sxs-lookup"><span data-stu-id="b9660-152">String</span></span>|<span data-ttu-id="b9660-153">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="b9660-153">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="b9660-154">String</span><span class="sxs-lookup"><span data-stu-id="b9660-154">String</span></span>|<span data-ttu-id="b9660-155">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="b9660-155">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b9660-156">要件</span><span class="sxs-lookup"><span data-stu-id="b9660-156">Requirements</span></span>

|<span data-ttu-id="b9660-157">要件</span><span class="sxs-lookup"><span data-stu-id="b9660-157">Requirement</span></span>| <span data-ttu-id="b9660-158">値</span><span class="sxs-lookup"><span data-stu-id="b9660-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9660-159">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b9660-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9660-160">1.0</span><span class="sxs-lookup"><span data-stu-id="b9660-160">1.0</span></span>|
|[<span data-ttu-id="b9660-161">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b9660-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9660-162">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b9660-162">Compose or Read</span></span>|

#### <a name="sourceproperty-string"></a><span data-ttu-id="b9660-163">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="b9660-163">SourceProperty: String</span></span>

<span data-ttu-id="b9660-164">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="b9660-164">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b9660-165">型</span><span class="sxs-lookup"><span data-stu-id="b9660-165">Type</span></span>

*   <span data-ttu-id="b9660-166">String</span><span class="sxs-lookup"><span data-stu-id="b9660-166">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b9660-167">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="b9660-167">Properties:</span></span>

|<span data-ttu-id="b9660-168">名前</span><span class="sxs-lookup"><span data-stu-id="b9660-168">Name</span></span>| <span data-ttu-id="b9660-169">種類</span><span class="sxs-lookup"><span data-stu-id="b9660-169">Type</span></span>| <span data-ttu-id="b9660-170">説明</span><span class="sxs-lookup"><span data-stu-id="b9660-170">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="b9660-171">String</span><span class="sxs-lookup"><span data-stu-id="b9660-171">String</span></span>|<span data-ttu-id="b9660-172">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="b9660-172">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="b9660-173">String</span><span class="sxs-lookup"><span data-stu-id="b9660-173">String</span></span>|<span data-ttu-id="b9660-174">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="b9660-174">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b9660-175">要件</span><span class="sxs-lookup"><span data-stu-id="b9660-175">Requirements</span></span>

|<span data-ttu-id="b9660-176">要件</span><span class="sxs-lookup"><span data-stu-id="b9660-176">Requirement</span></span>| <span data-ttu-id="b9660-177">値</span><span class="sxs-lookup"><span data-stu-id="b9660-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9660-178">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b9660-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9660-179">1.0</span><span class="sxs-lookup"><span data-stu-id="b9660-179">1.0</span></span>|
|[<span data-ttu-id="b9660-180">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b9660-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9660-181">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b9660-181">Compose or Read</span></span>|
