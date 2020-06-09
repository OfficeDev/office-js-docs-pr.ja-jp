---
title: Office 名前空間 - 要件セット 1.1
description: メールボックス API 要件セット1.1 を使用した Outlook アドインで使用可能な Office 名前空間メンバー。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: f127ab6594d5838700bbc04661d995b01da4f067
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611478"
---
# <a name="office-mailbox-requirement-set-11"></a><span data-ttu-id="acd71-103">Office (メールボックス要件セット 1.1)</span><span class="sxs-lookup"><span data-stu-id="acd71-103">Office (Mailbox requirement set 1.1)</span></span>

<span data-ttu-id="acd71-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="acd71-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="acd71-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="acd71-106">Requirements</span></span>

|<span data-ttu-id="acd71-107">要件</span><span class="sxs-lookup"><span data-stu-id="acd71-107">Requirement</span></span>| <span data-ttu-id="acd71-108">値</span><span class="sxs-lookup"><span data-stu-id="acd71-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="acd71-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="acd71-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="acd71-110">1.1</span><span class="sxs-lookup"><span data-stu-id="acd71-110">1.1</span></span>|
|[<span data-ttu-id="acd71-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="acd71-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="acd71-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="acd71-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="acd71-113">Properties</span><span class="sxs-lookup"><span data-stu-id="acd71-113">Properties</span></span>

| <span data-ttu-id="acd71-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="acd71-114">Property</span></span> | <span data-ttu-id="acd71-115">モード</span><span class="sxs-lookup"><span data-stu-id="acd71-115">Modes</span></span> | <span data-ttu-id="acd71-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="acd71-116">Return type</span></span> | <span data-ttu-id="acd71-117">最小値</span><span class="sxs-lookup"><span data-stu-id="acd71-117">Minimum</span></span><br><span data-ttu-id="acd71-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="acd71-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="acd71-119">context</span><span class="sxs-lookup"><span data-stu-id="acd71-119">context</span></span>](office.context.md) | <span data-ttu-id="acd71-120">作成</span><span class="sxs-lookup"><span data-stu-id="acd71-120">Compose</span></span><br><span data-ttu-id="acd71-121">Read</span><span class="sxs-lookup"><span data-stu-id="acd71-121">Read</span></span> | [<span data-ttu-id="acd71-122">Context</span><span class="sxs-lookup"><span data-stu-id="acd71-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.1) | [<span data-ttu-id="acd71-123">1.1</span><span class="sxs-lookup"><span data-stu-id="acd71-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="acd71-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="acd71-124">Enumerations</span></span>

| <span data-ttu-id="acd71-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="acd71-125">Enumeration</span></span> | <span data-ttu-id="acd71-126">モード</span><span class="sxs-lookup"><span data-stu-id="acd71-126">Modes</span></span> | <span data-ttu-id="acd71-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="acd71-127">Return type</span></span> | <span data-ttu-id="acd71-128">最小値</span><span class="sxs-lookup"><span data-stu-id="acd71-128">Minimum</span></span><br><span data-ttu-id="acd71-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="acd71-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="acd71-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="acd71-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="acd71-131">作成</span><span class="sxs-lookup"><span data-stu-id="acd71-131">Compose</span></span><br><span data-ttu-id="acd71-132">Read</span><span class="sxs-lookup"><span data-stu-id="acd71-132">Read</span></span> | <span data-ttu-id="acd71-133">String</span><span class="sxs-lookup"><span data-stu-id="acd71-133">String</span></span> | [<span data-ttu-id="acd71-134">1.1</span><span class="sxs-lookup"><span data-stu-id="acd71-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="acd71-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="acd71-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="acd71-136">作成</span><span class="sxs-lookup"><span data-stu-id="acd71-136">Compose</span></span><br><span data-ttu-id="acd71-137">Read</span><span class="sxs-lookup"><span data-stu-id="acd71-137">Read</span></span> | <span data-ttu-id="acd71-138">String</span><span class="sxs-lookup"><span data-stu-id="acd71-138">String</span></span> | [<span data-ttu-id="acd71-139">1.1</span><span class="sxs-lookup"><span data-stu-id="acd71-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="acd71-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="acd71-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="acd71-141">作成</span><span class="sxs-lookup"><span data-stu-id="acd71-141">Compose</span></span><br><span data-ttu-id="acd71-142">Read</span><span class="sxs-lookup"><span data-stu-id="acd71-142">Read</span></span> | <span data-ttu-id="acd71-143">String</span><span class="sxs-lookup"><span data-stu-id="acd71-143">String</span></span> | [<span data-ttu-id="acd71-144">1.1</span><span class="sxs-lookup"><span data-stu-id="acd71-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="acd71-145">名前空間</span><span class="sxs-lookup"><span data-stu-id="acd71-145">Namespaces</span></span>

<span data-ttu-id="acd71-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1):、、、、、など、多数の Outlook 固有の列挙を含み `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` ます。</span><span class="sxs-lookup"><span data-stu-id="acd71-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="acd71-147">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="acd71-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="acd71-148">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="acd71-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="acd71-149">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="acd71-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="acd71-150">型</span><span class="sxs-lookup"><span data-stu-id="acd71-150">Type</span></span>

*   <span data-ttu-id="acd71-151">String</span><span class="sxs-lookup"><span data-stu-id="acd71-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="acd71-152">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="acd71-152">Properties:</span></span>

|<span data-ttu-id="acd71-153">名前</span><span class="sxs-lookup"><span data-stu-id="acd71-153">Name</span></span>| <span data-ttu-id="acd71-154">種類</span><span class="sxs-lookup"><span data-stu-id="acd71-154">Type</span></span>| <span data-ttu-id="acd71-155">説明</span><span class="sxs-lookup"><span data-stu-id="acd71-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="acd71-156">String</span><span class="sxs-lookup"><span data-stu-id="acd71-156">String</span></span>|<span data-ttu-id="acd71-157">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="acd71-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="acd71-158">String</span><span class="sxs-lookup"><span data-stu-id="acd71-158">String</span></span>|<span data-ttu-id="acd71-159">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="acd71-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="acd71-160">Requirements</span><span class="sxs-lookup"><span data-stu-id="acd71-160">Requirements</span></span>

|<span data-ttu-id="acd71-161">要件</span><span class="sxs-lookup"><span data-stu-id="acd71-161">Requirement</span></span>| <span data-ttu-id="acd71-162">値</span><span class="sxs-lookup"><span data-stu-id="acd71-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="acd71-163">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="acd71-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="acd71-164">1.1</span><span class="sxs-lookup"><span data-stu-id="acd71-164">1.1</span></span>|
|[<span data-ttu-id="acd71-165">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="acd71-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="acd71-166">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="acd71-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="acd71-167">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="acd71-167">CoercionType: String</span></span>

<span data-ttu-id="acd71-168">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="acd71-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="acd71-169">型</span><span class="sxs-lookup"><span data-stu-id="acd71-169">Type</span></span>

*   <span data-ttu-id="acd71-170">String</span><span class="sxs-lookup"><span data-stu-id="acd71-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="acd71-171">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="acd71-171">Properties:</span></span>

|<span data-ttu-id="acd71-172">名前</span><span class="sxs-lookup"><span data-stu-id="acd71-172">Name</span></span>| <span data-ttu-id="acd71-173">種類</span><span class="sxs-lookup"><span data-stu-id="acd71-173">Type</span></span>| <span data-ttu-id="acd71-174">説明</span><span class="sxs-lookup"><span data-stu-id="acd71-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="acd71-175">String</span><span class="sxs-lookup"><span data-stu-id="acd71-175">String</span></span>|<span data-ttu-id="acd71-176">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="acd71-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="acd71-177">String</span><span class="sxs-lookup"><span data-stu-id="acd71-177">String</span></span>|<span data-ttu-id="acd71-178">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="acd71-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="acd71-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="acd71-179">Requirements</span></span>

|<span data-ttu-id="acd71-180">要件</span><span class="sxs-lookup"><span data-stu-id="acd71-180">Requirement</span></span>| <span data-ttu-id="acd71-181">値</span><span class="sxs-lookup"><span data-stu-id="acd71-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="acd71-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="acd71-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="acd71-183">1.1</span><span class="sxs-lookup"><span data-stu-id="acd71-183">1.1</span></span>|
|[<span data-ttu-id="acd71-184">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="acd71-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="acd71-185">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="acd71-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="acd71-186">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="acd71-186">SourceProperty: String</span></span>

<span data-ttu-id="acd71-187">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="acd71-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="acd71-188">型</span><span class="sxs-lookup"><span data-stu-id="acd71-188">Type</span></span>

*   <span data-ttu-id="acd71-189">String</span><span class="sxs-lookup"><span data-stu-id="acd71-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="acd71-190">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="acd71-190">Properties:</span></span>

|<span data-ttu-id="acd71-191">名前</span><span class="sxs-lookup"><span data-stu-id="acd71-191">Name</span></span>| <span data-ttu-id="acd71-192">種類</span><span class="sxs-lookup"><span data-stu-id="acd71-192">Type</span></span>| <span data-ttu-id="acd71-193">説明</span><span class="sxs-lookup"><span data-stu-id="acd71-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="acd71-194">String</span><span class="sxs-lookup"><span data-stu-id="acd71-194">String</span></span>|<span data-ttu-id="acd71-195">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="acd71-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="acd71-196">String</span><span class="sxs-lookup"><span data-stu-id="acd71-196">String</span></span>|<span data-ttu-id="acd71-197">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="acd71-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="acd71-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="acd71-198">Requirements</span></span>

|<span data-ttu-id="acd71-199">要件</span><span class="sxs-lookup"><span data-stu-id="acd71-199">Requirement</span></span>| <span data-ttu-id="acd71-200">値</span><span class="sxs-lookup"><span data-stu-id="acd71-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="acd71-201">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="acd71-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="acd71-202">1.1</span><span class="sxs-lookup"><span data-stu-id="acd71-202">1.1</span></span>|
|[<span data-ttu-id="acd71-203">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="acd71-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="acd71-204">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="acd71-204">Compose or Read</span></span>|