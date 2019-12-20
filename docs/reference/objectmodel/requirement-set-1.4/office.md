---
title: Office 名前空間-要件セット1.4
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: a2d3301448353ae3fbbc06be9f1fb2f7e1c3dfe6
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814858"
---
# <a name="office"></a><span data-ttu-id="bf4d2-102">Office</span><span class="sxs-lookup"><span data-stu-id="bf4d2-102">Office</span></span>

<span data-ttu-id="bf4d2-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bf4d2-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="bf4d2-105">要件</span><span class="sxs-lookup"><span data-stu-id="bf4d2-105">Requirements</span></span>

|<span data-ttu-id="bf4d2-106">要件</span><span class="sxs-lookup"><span data-stu-id="bf4d2-106">Requirement</span></span>| <span data-ttu-id="bf4d2-107">値</span><span class="sxs-lookup"><span data-stu-id="bf4d2-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf4d2-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bf4d2-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bf4d2-109">1.1</span><span class="sxs-lookup"><span data-stu-id="bf4d2-109">1.1</span></span>|
|[<span data-ttu-id="bf4d2-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bf4d2-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bf4d2-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bf4d2-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="bf4d2-112">Properties</span><span class="sxs-lookup"><span data-stu-id="bf4d2-112">Properties</span></span>

| <span data-ttu-id="bf4d2-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="bf4d2-113">Property</span></span> | <span data-ttu-id="bf4d2-114">モード</span><span class="sxs-lookup"><span data-stu-id="bf4d2-114">Modes</span></span> | <span data-ttu-id="bf4d2-115">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="bf4d2-115">Return type</span></span> | <span data-ttu-id="bf4d2-116">最小値</span><span class="sxs-lookup"><span data-stu-id="bf4d2-116">Minimum</span></span><br><span data-ttu-id="bf4d2-117">要件セット</span><span class="sxs-lookup"><span data-stu-id="bf4d2-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="bf4d2-118">context</span><span class="sxs-lookup"><span data-stu-id="bf4d2-118">context</span></span>](office.context.md) | <span data-ttu-id="bf4d2-119">作成</span><span class="sxs-lookup"><span data-stu-id="bf4d2-119">Compose</span></span><br><span data-ttu-id="bf4d2-120">読み取り</span><span class="sxs-lookup"><span data-stu-id="bf4d2-120">Read</span></span> | [<span data-ttu-id="bf4d2-121">Context</span><span class="sxs-lookup"><span data-stu-id="bf4d2-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.4) | [<span data-ttu-id="bf4d2-122">1.1</span><span class="sxs-lookup"><span data-stu-id="bf4d2-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="bf4d2-123">列挙型</span><span class="sxs-lookup"><span data-stu-id="bf4d2-123">Enumerations</span></span>

| <span data-ttu-id="bf4d2-124">列挙体</span><span class="sxs-lookup"><span data-stu-id="bf4d2-124">Enumeration</span></span> | <span data-ttu-id="bf4d2-125">モード</span><span class="sxs-lookup"><span data-stu-id="bf4d2-125">Modes</span></span> | <span data-ttu-id="bf4d2-126">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="bf4d2-126">Return type</span></span> | <span data-ttu-id="bf4d2-127">最小値</span><span class="sxs-lookup"><span data-stu-id="bf4d2-127">Minimum</span></span><br><span data-ttu-id="bf4d2-128">要件セット</span><span class="sxs-lookup"><span data-stu-id="bf4d2-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="bf4d2-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="bf4d2-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="bf4d2-130">作成</span><span class="sxs-lookup"><span data-stu-id="bf4d2-130">Compose</span></span><br><span data-ttu-id="bf4d2-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="bf4d2-131">Read</span></span> | <span data-ttu-id="bf4d2-132">String</span><span class="sxs-lookup"><span data-stu-id="bf4d2-132">String</span></span> | [<span data-ttu-id="bf4d2-133">1.1</span><span class="sxs-lookup"><span data-stu-id="bf4d2-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bf4d2-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="bf4d2-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="bf4d2-135">作成</span><span class="sxs-lookup"><span data-stu-id="bf4d2-135">Compose</span></span><br><span data-ttu-id="bf4d2-136">読み取り</span><span class="sxs-lookup"><span data-stu-id="bf4d2-136">Read</span></span> | <span data-ttu-id="bf4d2-137">String</span><span class="sxs-lookup"><span data-stu-id="bf4d2-137">String</span></span> | [<span data-ttu-id="bf4d2-138">1.1</span><span class="sxs-lookup"><span data-stu-id="bf4d2-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bf4d2-139">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="bf4d2-139">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="bf4d2-140">作成</span><span class="sxs-lookup"><span data-stu-id="bf4d2-140">Compose</span></span><br><span data-ttu-id="bf4d2-141">読み取り</span><span class="sxs-lookup"><span data-stu-id="bf4d2-141">Read</span></span> | <span data-ttu-id="bf4d2-142">String</span><span class="sxs-lookup"><span data-stu-id="bf4d2-142">String</span></span> | [<span data-ttu-id="bf4d2-143">1.1</span><span class="sxs-lookup"><span data-stu-id="bf4d2-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="bf4d2-144">名前空間</span><span class="sxs-lookup"><span data-stu-id="bf4d2-144">Namespaces</span></span>

<span data-ttu-id="bf4d2-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="bf4d2-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="bf4d2-146">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="bf4d2-146">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="bf4d2-147">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="bf4d2-147">AsyncResultStatus: String</span></span>

<span data-ttu-id="bf4d2-148">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="bf4d2-148">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="bf4d2-149">型</span><span class="sxs-lookup"><span data-stu-id="bf4d2-149">Type</span></span>

*   <span data-ttu-id="bf4d2-150">String</span><span class="sxs-lookup"><span data-stu-id="bf4d2-150">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bf4d2-151">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="bf4d2-151">Properties:</span></span>

|<span data-ttu-id="bf4d2-152">名前</span><span class="sxs-lookup"><span data-stu-id="bf4d2-152">Name</span></span>| <span data-ttu-id="bf4d2-153">種類</span><span class="sxs-lookup"><span data-stu-id="bf4d2-153">Type</span></span>| <span data-ttu-id="bf4d2-154">説明</span><span class="sxs-lookup"><span data-stu-id="bf4d2-154">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="bf4d2-155">String</span><span class="sxs-lookup"><span data-stu-id="bf4d2-155">String</span></span>|<span data-ttu-id="bf4d2-156">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="bf4d2-156">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="bf4d2-157">String</span><span class="sxs-lookup"><span data-stu-id="bf4d2-157">String</span></span>|<span data-ttu-id="bf4d2-158">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="bf4d2-158">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bf4d2-159">要件</span><span class="sxs-lookup"><span data-stu-id="bf4d2-159">Requirements</span></span>

|<span data-ttu-id="bf4d2-160">要件</span><span class="sxs-lookup"><span data-stu-id="bf4d2-160">Requirement</span></span>| <span data-ttu-id="bf4d2-161">値</span><span class="sxs-lookup"><span data-stu-id="bf4d2-161">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf4d2-162">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bf4d2-162">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bf4d2-163">1.1</span><span class="sxs-lookup"><span data-stu-id="bf4d2-163">1.1</span></span>|
|[<span data-ttu-id="bf4d2-164">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bf4d2-164">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bf4d2-165">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bf4d2-165">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="bf4d2-166">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="bf4d2-166">CoercionType: String</span></span>

<span data-ttu-id="bf4d2-167">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="bf4d2-167">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="bf4d2-168">型</span><span class="sxs-lookup"><span data-stu-id="bf4d2-168">Type</span></span>

*   <span data-ttu-id="bf4d2-169">String</span><span class="sxs-lookup"><span data-stu-id="bf4d2-169">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bf4d2-170">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="bf4d2-170">Properties:</span></span>

|<span data-ttu-id="bf4d2-171">名前</span><span class="sxs-lookup"><span data-stu-id="bf4d2-171">Name</span></span>| <span data-ttu-id="bf4d2-172">種類</span><span class="sxs-lookup"><span data-stu-id="bf4d2-172">Type</span></span>| <span data-ttu-id="bf4d2-173">説明</span><span class="sxs-lookup"><span data-stu-id="bf4d2-173">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="bf4d2-174">String</span><span class="sxs-lookup"><span data-stu-id="bf4d2-174">String</span></span>|<span data-ttu-id="bf4d2-175">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="bf4d2-175">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="bf4d2-176">String</span><span class="sxs-lookup"><span data-stu-id="bf4d2-176">String</span></span>|<span data-ttu-id="bf4d2-177">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="bf4d2-177">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bf4d2-178">要件</span><span class="sxs-lookup"><span data-stu-id="bf4d2-178">Requirements</span></span>

|<span data-ttu-id="bf4d2-179">要件</span><span class="sxs-lookup"><span data-stu-id="bf4d2-179">Requirement</span></span>| <span data-ttu-id="bf4d2-180">値</span><span class="sxs-lookup"><span data-stu-id="bf4d2-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf4d2-181">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bf4d2-181">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bf4d2-182">1.1</span><span class="sxs-lookup"><span data-stu-id="bf4d2-182">1.1</span></span>|
|[<span data-ttu-id="bf4d2-183">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bf4d2-183">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bf4d2-184">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bf4d2-184">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="bf4d2-185">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="bf4d2-185">SourceProperty: String</span></span>

<span data-ttu-id="bf4d2-186">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="bf4d2-186">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="bf4d2-187">型</span><span class="sxs-lookup"><span data-stu-id="bf4d2-187">Type</span></span>

*   <span data-ttu-id="bf4d2-188">String</span><span class="sxs-lookup"><span data-stu-id="bf4d2-188">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bf4d2-189">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="bf4d2-189">Properties:</span></span>

|<span data-ttu-id="bf4d2-190">名前</span><span class="sxs-lookup"><span data-stu-id="bf4d2-190">Name</span></span>| <span data-ttu-id="bf4d2-191">種類</span><span class="sxs-lookup"><span data-stu-id="bf4d2-191">Type</span></span>| <span data-ttu-id="bf4d2-192">説明</span><span class="sxs-lookup"><span data-stu-id="bf4d2-192">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="bf4d2-193">String</span><span class="sxs-lookup"><span data-stu-id="bf4d2-193">String</span></span>|<span data-ttu-id="bf4d2-194">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="bf4d2-194">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="bf4d2-195">String</span><span class="sxs-lookup"><span data-stu-id="bf4d2-195">String</span></span>|<span data-ttu-id="bf4d2-196">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="bf4d2-196">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bf4d2-197">要件</span><span class="sxs-lookup"><span data-stu-id="bf4d2-197">Requirements</span></span>

|<span data-ttu-id="bf4d2-198">要件</span><span class="sxs-lookup"><span data-stu-id="bf4d2-198">Requirement</span></span>| <span data-ttu-id="bf4d2-199">値</span><span class="sxs-lookup"><span data-stu-id="bf4d2-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="bf4d2-200">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bf4d2-200">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bf4d2-201">1.1</span><span class="sxs-lookup"><span data-stu-id="bf4d2-201">1.1</span></span>|
|[<span data-ttu-id="bf4d2-202">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bf4d2-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bf4d2-203">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bf4d2-203">Compose or Read</span></span>|
