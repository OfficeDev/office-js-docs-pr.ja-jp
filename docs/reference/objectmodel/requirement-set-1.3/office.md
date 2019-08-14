---
title: Office 名前空間-要件セット1.3
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: fb6036606f3a25cff5101351bd7df2b1b4cd2f21
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395632"
---
# <a name="office"></a><span data-ttu-id="a0eca-102">Office</span><span class="sxs-lookup"><span data-stu-id="a0eca-102">Office</span></span>

<span data-ttu-id="a0eca-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a0eca-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a0eca-105">要件</span><span class="sxs-lookup"><span data-stu-id="a0eca-105">Requirements</span></span>

|<span data-ttu-id="a0eca-106">要件</span><span class="sxs-lookup"><span data-stu-id="a0eca-106">Requirement</span></span>| <span data-ttu-id="a0eca-107">値</span><span class="sxs-lookup"><span data-stu-id="a0eca-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0eca-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a0eca-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0eca-109">1.0</span><span class="sxs-lookup"><span data-stu-id="a0eca-109">1.0</span></span>|
|[<span data-ttu-id="a0eca-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a0eca-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0eca-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="a0eca-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a0eca-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="a0eca-112">Members and methods</span></span>

| <span data-ttu-id="a0eca-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="a0eca-113">Member</span></span> | <span data-ttu-id="a0eca-114">種類</span><span class="sxs-lookup"><span data-stu-id="a0eca-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a0eca-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="a0eca-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="a0eca-116">Member</span><span class="sxs-lookup"><span data-stu-id="a0eca-116">Member</span></span> |
| [<span data-ttu-id="a0eca-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="a0eca-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="a0eca-118">Member</span><span class="sxs-lookup"><span data-stu-id="a0eca-118">Member</span></span> |
| [<span data-ttu-id="a0eca-119">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="a0eca-119">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="a0eca-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="a0eca-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="a0eca-121">名前空間</span><span class="sxs-lookup"><span data-stu-id="a0eca-121">Namespaces</span></span>

<span data-ttu-id="a0eca-122">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="a0eca-122">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="a0eca-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.3):、、、、、、などのさまざま`ItemType`な`EntityType`列挙`AttachmentType` `RecipientType` `ResponseType`値が含まれ`ItemNotificationMessageType`ています。</span><span class="sxs-lookup"><span data-stu-id="a0eca-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.3): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="a0eca-124">Members</span><span class="sxs-lookup"><span data-stu-id="a0eca-124">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="a0eca-125">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="a0eca-125">AsyncResultStatus: String</span></span>

<span data-ttu-id="a0eca-126">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="a0eca-126">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="a0eca-127">型</span><span class="sxs-lookup"><span data-stu-id="a0eca-127">Type</span></span>

*   <span data-ttu-id="a0eca-128">String</span><span class="sxs-lookup"><span data-stu-id="a0eca-128">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a0eca-129">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="a0eca-129">Properties:</span></span>

|<span data-ttu-id="a0eca-130">名前</span><span class="sxs-lookup"><span data-stu-id="a0eca-130">Name</span></span>| <span data-ttu-id="a0eca-131">種類</span><span class="sxs-lookup"><span data-stu-id="a0eca-131">Type</span></span>| <span data-ttu-id="a0eca-132">説明</span><span class="sxs-lookup"><span data-stu-id="a0eca-132">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="a0eca-133">String</span><span class="sxs-lookup"><span data-stu-id="a0eca-133">String</span></span>|<span data-ttu-id="a0eca-134">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="a0eca-134">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="a0eca-135">String</span><span class="sxs-lookup"><span data-stu-id="a0eca-135">String</span></span>|<span data-ttu-id="a0eca-136">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="a0eca-136">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0eca-137">要件</span><span class="sxs-lookup"><span data-stu-id="a0eca-137">Requirements</span></span>

|<span data-ttu-id="a0eca-138">要件</span><span class="sxs-lookup"><span data-stu-id="a0eca-138">Requirement</span></span>| <span data-ttu-id="a0eca-139">値</span><span class="sxs-lookup"><span data-stu-id="a0eca-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0eca-140">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a0eca-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0eca-141">1.0</span><span class="sxs-lookup"><span data-stu-id="a0eca-141">1.0</span></span>|
|[<span data-ttu-id="a0eca-142">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a0eca-142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0eca-143">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="a0eca-143">Compose or Read</span></span>|

#### <a name="coerciontype-string"></a><span data-ttu-id="a0eca-144">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="a0eca-144">CoercionType: String</span></span>

<span data-ttu-id="a0eca-145">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="a0eca-145">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a0eca-146">型</span><span class="sxs-lookup"><span data-stu-id="a0eca-146">Type</span></span>

*   <span data-ttu-id="a0eca-147">String</span><span class="sxs-lookup"><span data-stu-id="a0eca-147">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a0eca-148">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="a0eca-148">Properties:</span></span>

|<span data-ttu-id="a0eca-149">名前</span><span class="sxs-lookup"><span data-stu-id="a0eca-149">Name</span></span>| <span data-ttu-id="a0eca-150">種類</span><span class="sxs-lookup"><span data-stu-id="a0eca-150">Type</span></span>| <span data-ttu-id="a0eca-151">説明</span><span class="sxs-lookup"><span data-stu-id="a0eca-151">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="a0eca-152">String</span><span class="sxs-lookup"><span data-stu-id="a0eca-152">String</span></span>|<span data-ttu-id="a0eca-153">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="a0eca-153">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="a0eca-154">String</span><span class="sxs-lookup"><span data-stu-id="a0eca-154">String</span></span>|<span data-ttu-id="a0eca-155">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="a0eca-155">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0eca-156">要件</span><span class="sxs-lookup"><span data-stu-id="a0eca-156">Requirements</span></span>

|<span data-ttu-id="a0eca-157">要件</span><span class="sxs-lookup"><span data-stu-id="a0eca-157">Requirement</span></span>| <span data-ttu-id="a0eca-158">値</span><span class="sxs-lookup"><span data-stu-id="a0eca-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0eca-159">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a0eca-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0eca-160">1.0</span><span class="sxs-lookup"><span data-stu-id="a0eca-160">1.0</span></span>|
|[<span data-ttu-id="a0eca-161">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a0eca-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0eca-162">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="a0eca-162">Compose or Read</span></span>|

#### <a name="sourceproperty-string"></a><span data-ttu-id="a0eca-163">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="a0eca-163">SourceProperty: String</span></span>

<span data-ttu-id="a0eca-164">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="a0eca-164">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a0eca-165">型</span><span class="sxs-lookup"><span data-stu-id="a0eca-165">Type</span></span>

*   <span data-ttu-id="a0eca-166">String</span><span class="sxs-lookup"><span data-stu-id="a0eca-166">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a0eca-167">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="a0eca-167">Properties:</span></span>

|<span data-ttu-id="a0eca-168">名前</span><span class="sxs-lookup"><span data-stu-id="a0eca-168">Name</span></span>| <span data-ttu-id="a0eca-169">種類</span><span class="sxs-lookup"><span data-stu-id="a0eca-169">Type</span></span>| <span data-ttu-id="a0eca-170">説明</span><span class="sxs-lookup"><span data-stu-id="a0eca-170">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="a0eca-171">String</span><span class="sxs-lookup"><span data-stu-id="a0eca-171">String</span></span>|<span data-ttu-id="a0eca-172">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="a0eca-172">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="a0eca-173">String</span><span class="sxs-lookup"><span data-stu-id="a0eca-173">String</span></span>|<span data-ttu-id="a0eca-174">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="a0eca-174">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a0eca-175">要件</span><span class="sxs-lookup"><span data-stu-id="a0eca-175">Requirements</span></span>

|<span data-ttu-id="a0eca-176">要件</span><span class="sxs-lookup"><span data-stu-id="a0eca-176">Requirement</span></span>| <span data-ttu-id="a0eca-177">値</span><span class="sxs-lookup"><span data-stu-id="a0eca-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="a0eca-178">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a0eca-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a0eca-179">1.0</span><span class="sxs-lookup"><span data-stu-id="a0eca-179">1.0</span></span>|
|[<span data-ttu-id="a0eca-180">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a0eca-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a0eca-181">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="a0eca-181">Compose or Read</span></span>|
