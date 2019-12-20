---
title: Office 名前空間-要件セット1.2
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 714bbd6dfdcd47687a2309c24fd666168aed6556
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814328"
---
# <a name="office"></a><span data-ttu-id="08d16-102">Office</span><span class="sxs-lookup"><span data-stu-id="08d16-102">Office</span></span>

<span data-ttu-id="08d16-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="08d16-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="08d16-105">要件</span><span class="sxs-lookup"><span data-stu-id="08d16-105">Requirements</span></span>

|<span data-ttu-id="08d16-106">要件</span><span class="sxs-lookup"><span data-stu-id="08d16-106">Requirement</span></span>| <span data-ttu-id="08d16-107">値</span><span class="sxs-lookup"><span data-stu-id="08d16-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="08d16-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="08d16-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="08d16-109">1.1</span><span class="sxs-lookup"><span data-stu-id="08d16-109">1.1</span></span>|
|[<span data-ttu-id="08d16-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="08d16-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="08d16-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="08d16-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="08d16-112">Properties</span><span class="sxs-lookup"><span data-stu-id="08d16-112">Properties</span></span>

| <span data-ttu-id="08d16-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="08d16-113">Property</span></span> | <span data-ttu-id="08d16-114">モード</span><span class="sxs-lookup"><span data-stu-id="08d16-114">Modes</span></span> | <span data-ttu-id="08d16-115">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="08d16-115">Return type</span></span> | <span data-ttu-id="08d16-116">最小値</span><span class="sxs-lookup"><span data-stu-id="08d16-116">Minimum</span></span><br><span data-ttu-id="08d16-117">要件セット</span><span class="sxs-lookup"><span data-stu-id="08d16-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="08d16-118">context</span><span class="sxs-lookup"><span data-stu-id="08d16-118">context</span></span>](office.context.md) | <span data-ttu-id="08d16-119">作成</span><span class="sxs-lookup"><span data-stu-id="08d16-119">Compose</span></span><br><span data-ttu-id="08d16-120">読み取り</span><span class="sxs-lookup"><span data-stu-id="08d16-120">Read</span></span> | [<span data-ttu-id="08d16-121">Context</span><span class="sxs-lookup"><span data-stu-id="08d16-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.2) | [<span data-ttu-id="08d16-122">1.1</span><span class="sxs-lookup"><span data-stu-id="08d16-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="08d16-123">列挙型</span><span class="sxs-lookup"><span data-stu-id="08d16-123">Enumerations</span></span>

| <span data-ttu-id="08d16-124">列挙体</span><span class="sxs-lookup"><span data-stu-id="08d16-124">Enumeration</span></span> | <span data-ttu-id="08d16-125">モード</span><span class="sxs-lookup"><span data-stu-id="08d16-125">Modes</span></span> | <span data-ttu-id="08d16-126">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="08d16-126">Return type</span></span> | <span data-ttu-id="08d16-127">最小値</span><span class="sxs-lookup"><span data-stu-id="08d16-127">Minimum</span></span><br><span data-ttu-id="08d16-128">要件セット</span><span class="sxs-lookup"><span data-stu-id="08d16-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="08d16-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="08d16-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="08d16-130">作成</span><span class="sxs-lookup"><span data-stu-id="08d16-130">Compose</span></span><br><span data-ttu-id="08d16-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="08d16-131">Read</span></span> | <span data-ttu-id="08d16-132">String</span><span class="sxs-lookup"><span data-stu-id="08d16-132">String</span></span> | [<span data-ttu-id="08d16-133">1.1</span><span class="sxs-lookup"><span data-stu-id="08d16-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="08d16-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="08d16-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="08d16-135">作成</span><span class="sxs-lookup"><span data-stu-id="08d16-135">Compose</span></span><br><span data-ttu-id="08d16-136">読み取り</span><span class="sxs-lookup"><span data-stu-id="08d16-136">Read</span></span> | <span data-ttu-id="08d16-137">String</span><span class="sxs-lookup"><span data-stu-id="08d16-137">String</span></span> | [<span data-ttu-id="08d16-138">1.1</span><span class="sxs-lookup"><span data-stu-id="08d16-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="08d16-139">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="08d16-139">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="08d16-140">作成</span><span class="sxs-lookup"><span data-stu-id="08d16-140">Compose</span></span><br><span data-ttu-id="08d16-141">読み取り</span><span class="sxs-lookup"><span data-stu-id="08d16-141">Read</span></span> | <span data-ttu-id="08d16-142">String</span><span class="sxs-lookup"><span data-stu-id="08d16-142">String</span></span> | [<span data-ttu-id="08d16-143">1.1</span><span class="sxs-lookup"><span data-stu-id="08d16-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="08d16-144">名前空間</span><span class="sxs-lookup"><span data-stu-id="08d16-144">Namespaces</span></span>

<span data-ttu-id="08d16-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="08d16-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="08d16-146">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="08d16-146">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="08d16-147">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="08d16-147">AsyncResultStatus: String</span></span>

<span data-ttu-id="08d16-148">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="08d16-148">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="08d16-149">型</span><span class="sxs-lookup"><span data-stu-id="08d16-149">Type</span></span>

*   <span data-ttu-id="08d16-150">String</span><span class="sxs-lookup"><span data-stu-id="08d16-150">String</span></span>

##### <a name="properties"></a><span data-ttu-id="08d16-151">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="08d16-151">Properties:</span></span>

|<span data-ttu-id="08d16-152">名前</span><span class="sxs-lookup"><span data-stu-id="08d16-152">Name</span></span>| <span data-ttu-id="08d16-153">種類</span><span class="sxs-lookup"><span data-stu-id="08d16-153">Type</span></span>| <span data-ttu-id="08d16-154">説明</span><span class="sxs-lookup"><span data-stu-id="08d16-154">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="08d16-155">String</span><span class="sxs-lookup"><span data-stu-id="08d16-155">String</span></span>|<span data-ttu-id="08d16-156">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="08d16-156">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="08d16-157">String</span><span class="sxs-lookup"><span data-stu-id="08d16-157">String</span></span>|<span data-ttu-id="08d16-158">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="08d16-158">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="08d16-159">要件</span><span class="sxs-lookup"><span data-stu-id="08d16-159">Requirements</span></span>

|<span data-ttu-id="08d16-160">要件</span><span class="sxs-lookup"><span data-stu-id="08d16-160">Requirement</span></span>| <span data-ttu-id="08d16-161">値</span><span class="sxs-lookup"><span data-stu-id="08d16-161">Value</span></span>|
|---|---|
|[<span data-ttu-id="08d16-162">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="08d16-162">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="08d16-163">1.1</span><span class="sxs-lookup"><span data-stu-id="08d16-163">1.1</span></span>|
|[<span data-ttu-id="08d16-164">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="08d16-164">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="08d16-165">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="08d16-165">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="08d16-166">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="08d16-166">CoercionType: String</span></span>

<span data-ttu-id="08d16-167">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="08d16-167">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="08d16-168">型</span><span class="sxs-lookup"><span data-stu-id="08d16-168">Type</span></span>

*   <span data-ttu-id="08d16-169">String</span><span class="sxs-lookup"><span data-stu-id="08d16-169">String</span></span>

##### <a name="properties"></a><span data-ttu-id="08d16-170">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="08d16-170">Properties:</span></span>

|<span data-ttu-id="08d16-171">名前</span><span class="sxs-lookup"><span data-stu-id="08d16-171">Name</span></span>| <span data-ttu-id="08d16-172">種類</span><span class="sxs-lookup"><span data-stu-id="08d16-172">Type</span></span>| <span data-ttu-id="08d16-173">説明</span><span class="sxs-lookup"><span data-stu-id="08d16-173">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="08d16-174">String</span><span class="sxs-lookup"><span data-stu-id="08d16-174">String</span></span>|<span data-ttu-id="08d16-175">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="08d16-175">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="08d16-176">String</span><span class="sxs-lookup"><span data-stu-id="08d16-176">String</span></span>|<span data-ttu-id="08d16-177">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="08d16-177">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="08d16-178">要件</span><span class="sxs-lookup"><span data-stu-id="08d16-178">Requirements</span></span>

|<span data-ttu-id="08d16-179">要件</span><span class="sxs-lookup"><span data-stu-id="08d16-179">Requirement</span></span>| <span data-ttu-id="08d16-180">値</span><span class="sxs-lookup"><span data-stu-id="08d16-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="08d16-181">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="08d16-181">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="08d16-182">1.1</span><span class="sxs-lookup"><span data-stu-id="08d16-182">1.1</span></span>|
|[<span data-ttu-id="08d16-183">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="08d16-183">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="08d16-184">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="08d16-184">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="08d16-185">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="08d16-185">SourceProperty: String</span></span>

<span data-ttu-id="08d16-186">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="08d16-186">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="08d16-187">型</span><span class="sxs-lookup"><span data-stu-id="08d16-187">Type</span></span>

*   <span data-ttu-id="08d16-188">String</span><span class="sxs-lookup"><span data-stu-id="08d16-188">String</span></span>

##### <a name="properties"></a><span data-ttu-id="08d16-189">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="08d16-189">Properties:</span></span>

|<span data-ttu-id="08d16-190">名前</span><span class="sxs-lookup"><span data-stu-id="08d16-190">Name</span></span>| <span data-ttu-id="08d16-191">種類</span><span class="sxs-lookup"><span data-stu-id="08d16-191">Type</span></span>| <span data-ttu-id="08d16-192">説明</span><span class="sxs-lookup"><span data-stu-id="08d16-192">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="08d16-193">String</span><span class="sxs-lookup"><span data-stu-id="08d16-193">String</span></span>|<span data-ttu-id="08d16-194">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="08d16-194">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="08d16-195">String</span><span class="sxs-lookup"><span data-stu-id="08d16-195">String</span></span>|<span data-ttu-id="08d16-196">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="08d16-196">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="08d16-197">要件</span><span class="sxs-lookup"><span data-stu-id="08d16-197">Requirements</span></span>

|<span data-ttu-id="08d16-198">要件</span><span class="sxs-lookup"><span data-stu-id="08d16-198">Requirement</span></span>| <span data-ttu-id="08d16-199">値</span><span class="sxs-lookup"><span data-stu-id="08d16-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="08d16-200">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="08d16-200">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="08d16-201">1.1</span><span class="sxs-lookup"><span data-stu-id="08d16-201">1.1</span></span>|
|[<span data-ttu-id="08d16-202">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="08d16-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="08d16-203">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="08d16-203">Compose or Read</span></span>|
