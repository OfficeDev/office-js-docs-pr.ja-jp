---
title: Office名前空間 - 要件セット 1.4
description: Office API 要件セット 1.4 をOutlookアドインで使用できる名前空間メンバーを指定します。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 0221ab09048719317c131f0204e2fc60c4f8f7d4
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591024"
---
# <a name="office-mailbox-requirement-set-14"></a><span data-ttu-id="7d862-103">Office (メールボックス要件セット 1.4)</span><span class="sxs-lookup"><span data-stu-id="7d862-103">Office (Mailbox requirement set 1.4)</span></span>

<span data-ttu-id="7d862-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7d862-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="7d862-106">要件</span><span class="sxs-lookup"><span data-stu-id="7d862-106">Requirements</span></span>

|<span data-ttu-id="7d862-107">要件</span><span class="sxs-lookup"><span data-stu-id="7d862-107">Requirement</span></span>| <span data-ttu-id="7d862-108">値</span><span class="sxs-lookup"><span data-stu-id="7d862-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="7d862-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7d862-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7d862-110">1.1</span><span class="sxs-lookup"><span data-stu-id="7d862-110">1.1</span></span>|
|[<span data-ttu-id="7d862-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7d862-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7d862-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7d862-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="7d862-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="7d862-113">Properties</span></span>

| <span data-ttu-id="7d862-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="7d862-114">Property</span></span> | <span data-ttu-id="7d862-115">モード</span><span class="sxs-lookup"><span data-stu-id="7d862-115">Modes</span></span> | <span data-ttu-id="7d862-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="7d862-116">Return type</span></span> | <span data-ttu-id="7d862-117">最小値</span><span class="sxs-lookup"><span data-stu-id="7d862-117">Minimum</span></span><br><span data-ttu-id="7d862-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="7d862-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="7d862-119">context</span><span class="sxs-lookup"><span data-stu-id="7d862-119">context</span></span>](office.context.md) | <span data-ttu-id="7d862-120">作成</span><span class="sxs-lookup"><span data-stu-id="7d862-120">Compose</span></span><br><span data-ttu-id="7d862-121">Read</span><span class="sxs-lookup"><span data-stu-id="7d862-121">Read</span></span> | [<span data-ttu-id="7d862-122">Context</span><span class="sxs-lookup"><span data-stu-id="7d862-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="7d862-123">1.1</span><span class="sxs-lookup"><span data-stu-id="7d862-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="7d862-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="7d862-124">Enumerations</span></span>

| <span data-ttu-id="7d862-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="7d862-125">Enumeration</span></span> | <span data-ttu-id="7d862-126">モード</span><span class="sxs-lookup"><span data-stu-id="7d862-126">Modes</span></span> | <span data-ttu-id="7d862-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="7d862-127">Return type</span></span> | <span data-ttu-id="7d862-128">最小値</span><span class="sxs-lookup"><span data-stu-id="7d862-128">Minimum</span></span><br><span data-ttu-id="7d862-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="7d862-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="7d862-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="7d862-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="7d862-131">作成</span><span class="sxs-lookup"><span data-stu-id="7d862-131">Compose</span></span><br><span data-ttu-id="7d862-132">Read</span><span class="sxs-lookup"><span data-stu-id="7d862-132">Read</span></span> | <span data-ttu-id="7d862-133">String</span><span class="sxs-lookup"><span data-stu-id="7d862-133">String</span></span> | [<span data-ttu-id="7d862-134">1.1</span><span class="sxs-lookup"><span data-stu-id="7d862-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7d862-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="7d862-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="7d862-136">作成</span><span class="sxs-lookup"><span data-stu-id="7d862-136">Compose</span></span><br><span data-ttu-id="7d862-137">Read</span><span class="sxs-lookup"><span data-stu-id="7d862-137">Read</span></span> | <span data-ttu-id="7d862-138">String</span><span class="sxs-lookup"><span data-stu-id="7d862-138">String</span></span> | [<span data-ttu-id="7d862-139">1.1</span><span class="sxs-lookup"><span data-stu-id="7d862-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7d862-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="7d862-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="7d862-141">作成</span><span class="sxs-lookup"><span data-stu-id="7d862-141">Compose</span></span><br><span data-ttu-id="7d862-142">Read</span><span class="sxs-lookup"><span data-stu-id="7d862-142">Read</span></span> | <span data-ttu-id="7d862-143">String</span><span class="sxs-lookup"><span data-stu-id="7d862-143">String</span></span> | [<span data-ttu-id="7d862-144">1.1</span><span class="sxs-lookup"><span data-stu-id="7d862-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="7d862-145">名前空間</span><span class="sxs-lookup"><span data-stu-id="7d862-145">Namespaces</span></span>

<span data-ttu-id="7d862-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4&preserve-view=true): 、 など、Outlook固有の列挙の `ItemType` `EntityType` `AttachmentType` 数 `RecipientType` が `ResponseType` 含まれています `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="7d862-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="7d862-147">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="7d862-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="7d862-148">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="7d862-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="7d862-149">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="7d862-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="7d862-150">型</span><span class="sxs-lookup"><span data-stu-id="7d862-150">Type</span></span>

*   <span data-ttu-id="7d862-151">String</span><span class="sxs-lookup"><span data-stu-id="7d862-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7d862-152">プロパティ</span><span class="sxs-lookup"><span data-stu-id="7d862-152">Properties</span></span>

|<span data-ttu-id="7d862-153">名前</span><span class="sxs-lookup"><span data-stu-id="7d862-153">Name</span></span>| <span data-ttu-id="7d862-154">型</span><span class="sxs-lookup"><span data-stu-id="7d862-154">Type</span></span>| <span data-ttu-id="7d862-155">説明</span><span class="sxs-lookup"><span data-stu-id="7d862-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="7d862-156">String</span><span class="sxs-lookup"><span data-stu-id="7d862-156">String</span></span>|<span data-ttu-id="7d862-157">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="7d862-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="7d862-158">String</span><span class="sxs-lookup"><span data-stu-id="7d862-158">String</span></span>|<span data-ttu-id="7d862-159">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="7d862-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7d862-160">要件</span><span class="sxs-lookup"><span data-stu-id="7d862-160">Requirements</span></span>

|<span data-ttu-id="7d862-161">要件</span><span class="sxs-lookup"><span data-stu-id="7d862-161">Requirement</span></span>| <span data-ttu-id="7d862-162">値</span><span class="sxs-lookup"><span data-stu-id="7d862-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="7d862-163">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7d862-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7d862-164">1.1</span><span class="sxs-lookup"><span data-stu-id="7d862-164">1.1</span></span>|
|[<span data-ttu-id="7d862-165">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7d862-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7d862-166">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7d862-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="7d862-167">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="7d862-167">CoercionType: String</span></span>

<span data-ttu-id="7d862-168">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="7d862-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="7d862-169">型</span><span class="sxs-lookup"><span data-stu-id="7d862-169">Type</span></span>

*   <span data-ttu-id="7d862-170">String</span><span class="sxs-lookup"><span data-stu-id="7d862-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7d862-171">プロパティ</span><span class="sxs-lookup"><span data-stu-id="7d862-171">Properties</span></span>

|<span data-ttu-id="7d862-172">名前</span><span class="sxs-lookup"><span data-stu-id="7d862-172">Name</span></span>| <span data-ttu-id="7d862-173">型</span><span class="sxs-lookup"><span data-stu-id="7d862-173">Type</span></span>| <span data-ttu-id="7d862-174">説明</span><span class="sxs-lookup"><span data-stu-id="7d862-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="7d862-175">String</span><span class="sxs-lookup"><span data-stu-id="7d862-175">String</span></span>|<span data-ttu-id="7d862-176">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="7d862-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="7d862-177">String</span><span class="sxs-lookup"><span data-stu-id="7d862-177">String</span></span>|<span data-ttu-id="7d862-178">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="7d862-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7d862-179">要件</span><span class="sxs-lookup"><span data-stu-id="7d862-179">Requirements</span></span>

|<span data-ttu-id="7d862-180">要件</span><span class="sxs-lookup"><span data-stu-id="7d862-180">Requirement</span></span>| <span data-ttu-id="7d862-181">値</span><span class="sxs-lookup"><span data-stu-id="7d862-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="7d862-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7d862-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7d862-183">1.1</span><span class="sxs-lookup"><span data-stu-id="7d862-183">1.1</span></span>|
|[<span data-ttu-id="7d862-184">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7d862-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7d862-185">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7d862-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="7d862-186">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="7d862-186">SourceProperty: String</span></span>

<span data-ttu-id="7d862-187">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="7d862-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="7d862-188">型</span><span class="sxs-lookup"><span data-stu-id="7d862-188">Type</span></span>

*   <span data-ttu-id="7d862-189">String</span><span class="sxs-lookup"><span data-stu-id="7d862-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7d862-190">プロパティ</span><span class="sxs-lookup"><span data-stu-id="7d862-190">Properties</span></span>

|<span data-ttu-id="7d862-191">名前</span><span class="sxs-lookup"><span data-stu-id="7d862-191">Name</span></span>| <span data-ttu-id="7d862-192">型</span><span class="sxs-lookup"><span data-stu-id="7d862-192">Type</span></span>| <span data-ttu-id="7d862-193">説明</span><span class="sxs-lookup"><span data-stu-id="7d862-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="7d862-194">String</span><span class="sxs-lookup"><span data-stu-id="7d862-194">String</span></span>|<span data-ttu-id="7d862-195">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="7d862-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="7d862-196">String</span><span class="sxs-lookup"><span data-stu-id="7d862-196">String</span></span>|<span data-ttu-id="7d862-197">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="7d862-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7d862-198">要件</span><span class="sxs-lookup"><span data-stu-id="7d862-198">Requirements</span></span>

|<span data-ttu-id="7d862-199">要件</span><span class="sxs-lookup"><span data-stu-id="7d862-199">Requirement</span></span>| <span data-ttu-id="7d862-200">値</span><span class="sxs-lookup"><span data-stu-id="7d862-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="7d862-201">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7d862-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7d862-202">1.1</span><span class="sxs-lookup"><span data-stu-id="7d862-202">1.1</span></span>|
|[<span data-ttu-id="7d862-203">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7d862-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7d862-204">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7d862-204">Compose or Read</span></span>|
