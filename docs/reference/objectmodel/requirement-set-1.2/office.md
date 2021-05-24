---
title: Office名前空間 - 要件セット 1.2
description: Office API 要件セット 1.2 を使用Outlookアドインで使用できる名前空間メンバーを指定します。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 4cd15d77d1c5d9b95152f038f3421c5838bfb84f
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590408"
---
# <a name="office-mailbox-requirement-set-12"></a><span data-ttu-id="9c272-103">Office (メールボックス要件セット 1.2)</span><span class="sxs-lookup"><span data-stu-id="9c272-103">Office (Mailbox requirement set 1.2)</span></span>

<span data-ttu-id="9c272-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9c272-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c272-106">要件</span><span class="sxs-lookup"><span data-stu-id="9c272-106">Requirements</span></span>

|<span data-ttu-id="9c272-107">要件</span><span class="sxs-lookup"><span data-stu-id="9c272-107">Requirement</span></span>| <span data-ttu-id="9c272-108">値</span><span class="sxs-lookup"><span data-stu-id="9c272-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c272-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9c272-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9c272-110">1.1</span><span class="sxs-lookup"><span data-stu-id="9c272-110">1.1</span></span>|
|[<span data-ttu-id="9c272-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9c272-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9c272-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9c272-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="9c272-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="9c272-113">Properties</span></span>

| <span data-ttu-id="9c272-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="9c272-114">Property</span></span> | <span data-ttu-id="9c272-115">モード</span><span class="sxs-lookup"><span data-stu-id="9c272-115">Modes</span></span> | <span data-ttu-id="9c272-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="9c272-116">Return type</span></span> | <span data-ttu-id="9c272-117">最小値</span><span class="sxs-lookup"><span data-stu-id="9c272-117">Minimum</span></span><br><span data-ttu-id="9c272-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="9c272-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="9c272-119">context</span><span class="sxs-lookup"><span data-stu-id="9c272-119">context</span></span>](office.context.md) | <span data-ttu-id="9c272-120">作成</span><span class="sxs-lookup"><span data-stu-id="9c272-120">Compose</span></span><br><span data-ttu-id="9c272-121">Read</span><span class="sxs-lookup"><span data-stu-id="9c272-121">Read</span></span> | [<span data-ttu-id="9c272-122">Context</span><span class="sxs-lookup"><span data-stu-id="9c272-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="9c272-123">1.1</span><span class="sxs-lookup"><span data-stu-id="9c272-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="9c272-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="9c272-124">Enumerations</span></span>

| <span data-ttu-id="9c272-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="9c272-125">Enumeration</span></span> | <span data-ttu-id="9c272-126">モード</span><span class="sxs-lookup"><span data-stu-id="9c272-126">Modes</span></span> | <span data-ttu-id="9c272-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="9c272-127">Return type</span></span> | <span data-ttu-id="9c272-128">最小値</span><span class="sxs-lookup"><span data-stu-id="9c272-128">Minimum</span></span><br><span data-ttu-id="9c272-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="9c272-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="9c272-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="9c272-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="9c272-131">作成</span><span class="sxs-lookup"><span data-stu-id="9c272-131">Compose</span></span><br><span data-ttu-id="9c272-132">Read</span><span class="sxs-lookup"><span data-stu-id="9c272-132">Read</span></span> | <span data-ttu-id="9c272-133">String</span><span class="sxs-lookup"><span data-stu-id="9c272-133">String</span></span> | [<span data-ttu-id="9c272-134">1.1</span><span class="sxs-lookup"><span data-stu-id="9c272-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9c272-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="9c272-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="9c272-136">作成</span><span class="sxs-lookup"><span data-stu-id="9c272-136">Compose</span></span><br><span data-ttu-id="9c272-137">Read</span><span class="sxs-lookup"><span data-stu-id="9c272-137">Read</span></span> | <span data-ttu-id="9c272-138">String</span><span class="sxs-lookup"><span data-stu-id="9c272-138">String</span></span> | [<span data-ttu-id="9c272-139">1.1</span><span class="sxs-lookup"><span data-stu-id="9c272-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9c272-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="9c272-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="9c272-141">作成</span><span class="sxs-lookup"><span data-stu-id="9c272-141">Compose</span></span><br><span data-ttu-id="9c272-142">Read</span><span class="sxs-lookup"><span data-stu-id="9c272-142">Read</span></span> | <span data-ttu-id="9c272-143">String</span><span class="sxs-lookup"><span data-stu-id="9c272-143">String</span></span> | [<span data-ttu-id="9c272-144">1.1</span><span class="sxs-lookup"><span data-stu-id="9c272-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="9c272-145">名前空間</span><span class="sxs-lookup"><span data-stu-id="9c272-145">Namespaces</span></span>

<span data-ttu-id="9c272-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2&preserve-view=true): 、 など、Outlook固有の列挙の `ItemType` `EntityType` `AttachmentType` 数 `RecipientType` が `ResponseType` 含まれています `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="9c272-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="9c272-147">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="9c272-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="9c272-148">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="9c272-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="9c272-149">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="9c272-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="9c272-150">型</span><span class="sxs-lookup"><span data-stu-id="9c272-150">Type</span></span>

*   <span data-ttu-id="9c272-151">String</span><span class="sxs-lookup"><span data-stu-id="9c272-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9c272-152">プロパティ</span><span class="sxs-lookup"><span data-stu-id="9c272-152">Properties</span></span>

|<span data-ttu-id="9c272-153">名前</span><span class="sxs-lookup"><span data-stu-id="9c272-153">Name</span></span>| <span data-ttu-id="9c272-154">型</span><span class="sxs-lookup"><span data-stu-id="9c272-154">Type</span></span>| <span data-ttu-id="9c272-155">説明</span><span class="sxs-lookup"><span data-stu-id="9c272-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="9c272-156">String</span><span class="sxs-lookup"><span data-stu-id="9c272-156">String</span></span>|<span data-ttu-id="9c272-157">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="9c272-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="9c272-158">String</span><span class="sxs-lookup"><span data-stu-id="9c272-158">String</span></span>|<span data-ttu-id="9c272-159">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="9c272-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9c272-160">要件</span><span class="sxs-lookup"><span data-stu-id="9c272-160">Requirements</span></span>

|<span data-ttu-id="9c272-161">要件</span><span class="sxs-lookup"><span data-stu-id="9c272-161">Requirement</span></span>| <span data-ttu-id="9c272-162">値</span><span class="sxs-lookup"><span data-stu-id="9c272-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c272-163">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9c272-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9c272-164">1.1</span><span class="sxs-lookup"><span data-stu-id="9c272-164">1.1</span></span>|
|[<span data-ttu-id="9c272-165">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9c272-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9c272-166">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9c272-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="9c272-167">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="9c272-167">CoercionType: String</span></span>

<span data-ttu-id="9c272-168">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="9c272-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9c272-169">型</span><span class="sxs-lookup"><span data-stu-id="9c272-169">Type</span></span>

*   <span data-ttu-id="9c272-170">String</span><span class="sxs-lookup"><span data-stu-id="9c272-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9c272-171">プロパティ</span><span class="sxs-lookup"><span data-stu-id="9c272-171">Properties</span></span>

|<span data-ttu-id="9c272-172">名前</span><span class="sxs-lookup"><span data-stu-id="9c272-172">Name</span></span>| <span data-ttu-id="9c272-173">型</span><span class="sxs-lookup"><span data-stu-id="9c272-173">Type</span></span>| <span data-ttu-id="9c272-174">説明</span><span class="sxs-lookup"><span data-stu-id="9c272-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="9c272-175">String</span><span class="sxs-lookup"><span data-stu-id="9c272-175">String</span></span>|<span data-ttu-id="9c272-176">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="9c272-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="9c272-177">String</span><span class="sxs-lookup"><span data-stu-id="9c272-177">String</span></span>|<span data-ttu-id="9c272-178">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="9c272-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9c272-179">要件</span><span class="sxs-lookup"><span data-stu-id="9c272-179">Requirements</span></span>

|<span data-ttu-id="9c272-180">要件</span><span class="sxs-lookup"><span data-stu-id="9c272-180">Requirement</span></span>| <span data-ttu-id="9c272-181">値</span><span class="sxs-lookup"><span data-stu-id="9c272-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c272-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9c272-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9c272-183">1.1</span><span class="sxs-lookup"><span data-stu-id="9c272-183">1.1</span></span>|
|[<span data-ttu-id="9c272-184">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9c272-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9c272-185">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9c272-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="9c272-186">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="9c272-186">SourceProperty: String</span></span>

<span data-ttu-id="9c272-187">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="9c272-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9c272-188">型</span><span class="sxs-lookup"><span data-stu-id="9c272-188">Type</span></span>

*   <span data-ttu-id="9c272-189">String</span><span class="sxs-lookup"><span data-stu-id="9c272-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9c272-190">プロパティ</span><span class="sxs-lookup"><span data-stu-id="9c272-190">Properties</span></span>

|<span data-ttu-id="9c272-191">名前</span><span class="sxs-lookup"><span data-stu-id="9c272-191">Name</span></span>| <span data-ttu-id="9c272-192">型</span><span class="sxs-lookup"><span data-stu-id="9c272-192">Type</span></span>| <span data-ttu-id="9c272-193">説明</span><span class="sxs-lookup"><span data-stu-id="9c272-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="9c272-194">String</span><span class="sxs-lookup"><span data-stu-id="9c272-194">String</span></span>|<span data-ttu-id="9c272-195">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="9c272-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="9c272-196">String</span><span class="sxs-lookup"><span data-stu-id="9c272-196">String</span></span>|<span data-ttu-id="9c272-197">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="9c272-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9c272-198">要件</span><span class="sxs-lookup"><span data-stu-id="9c272-198">Requirements</span></span>

|<span data-ttu-id="9c272-199">要件</span><span class="sxs-lookup"><span data-stu-id="9c272-199">Requirement</span></span>| <span data-ttu-id="9c272-200">値</span><span class="sxs-lookup"><span data-stu-id="9c272-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c272-201">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="9c272-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9c272-202">1.1</span><span class="sxs-lookup"><span data-stu-id="9c272-202">1.1</span></span>|
|[<span data-ttu-id="9c272-203">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9c272-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9c272-204">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="9c272-204">Compose or Read</span></span>|
