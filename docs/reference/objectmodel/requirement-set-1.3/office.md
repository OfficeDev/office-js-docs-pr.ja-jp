---
title: Office 名前空間-要件セット1.3
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 3c6ddc34001f4d1622bc76d9bca1fbde9425be8b
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814900"
---
# <a name="office"></a><span data-ttu-id="8f75c-102">Office</span><span class="sxs-lookup"><span data-stu-id="8f75c-102">Office</span></span>

<span data-ttu-id="8f75c-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8f75c-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8f75c-105">要件</span><span class="sxs-lookup"><span data-stu-id="8f75c-105">Requirements</span></span>

|<span data-ttu-id="8f75c-106">要件</span><span class="sxs-lookup"><span data-stu-id="8f75c-106">Requirement</span></span>| <span data-ttu-id="8f75c-107">値</span><span class="sxs-lookup"><span data-stu-id="8f75c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f75c-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8f75c-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f75c-109">1.1</span><span class="sxs-lookup"><span data-stu-id="8f75c-109">1.1</span></span>|
|[<span data-ttu-id="8f75c-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8f75c-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8f75c-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8f75c-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="8f75c-112">Properties</span><span class="sxs-lookup"><span data-stu-id="8f75c-112">Properties</span></span>

| <span data-ttu-id="8f75c-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="8f75c-113">Property</span></span> | <span data-ttu-id="8f75c-114">モード</span><span class="sxs-lookup"><span data-stu-id="8f75c-114">Modes</span></span> | <span data-ttu-id="8f75c-115">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="8f75c-115">Return type</span></span> | <span data-ttu-id="8f75c-116">最小値</span><span class="sxs-lookup"><span data-stu-id="8f75c-116">Minimum</span></span><br><span data-ttu-id="8f75c-117">要件セット</span><span class="sxs-lookup"><span data-stu-id="8f75c-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8f75c-118">context</span><span class="sxs-lookup"><span data-stu-id="8f75c-118">context</span></span>](office.context.md) | <span data-ttu-id="8f75c-119">作成</span><span class="sxs-lookup"><span data-stu-id="8f75c-119">Compose</span></span><br><span data-ttu-id="8f75c-120">読み取り</span><span class="sxs-lookup"><span data-stu-id="8f75c-120">Read</span></span> | [<span data-ttu-id="8f75c-121">Context</span><span class="sxs-lookup"><span data-stu-id="8f75c-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.3) | [<span data-ttu-id="8f75c-122">1.1</span><span class="sxs-lookup"><span data-stu-id="8f75c-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="8f75c-123">列挙型</span><span class="sxs-lookup"><span data-stu-id="8f75c-123">Enumerations</span></span>

| <span data-ttu-id="8f75c-124">列挙体</span><span class="sxs-lookup"><span data-stu-id="8f75c-124">Enumeration</span></span> | <span data-ttu-id="8f75c-125">モード</span><span class="sxs-lookup"><span data-stu-id="8f75c-125">Modes</span></span> | <span data-ttu-id="8f75c-126">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="8f75c-126">Return type</span></span> | <span data-ttu-id="8f75c-127">最小値</span><span class="sxs-lookup"><span data-stu-id="8f75c-127">Minimum</span></span><br><span data-ttu-id="8f75c-128">要件セット</span><span class="sxs-lookup"><span data-stu-id="8f75c-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8f75c-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="8f75c-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="8f75c-130">作成</span><span class="sxs-lookup"><span data-stu-id="8f75c-130">Compose</span></span><br><span data-ttu-id="8f75c-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="8f75c-131">Read</span></span> | <span data-ttu-id="8f75c-132">String</span><span class="sxs-lookup"><span data-stu-id="8f75c-132">String</span></span> | [<span data-ttu-id="8f75c-133">1.1</span><span class="sxs-lookup"><span data-stu-id="8f75c-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8f75c-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="8f75c-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="8f75c-135">作成</span><span class="sxs-lookup"><span data-stu-id="8f75c-135">Compose</span></span><br><span data-ttu-id="8f75c-136">読み取り</span><span class="sxs-lookup"><span data-stu-id="8f75c-136">Read</span></span> | <span data-ttu-id="8f75c-137">String</span><span class="sxs-lookup"><span data-stu-id="8f75c-137">String</span></span> | [<span data-ttu-id="8f75c-138">1.1</span><span class="sxs-lookup"><span data-stu-id="8f75c-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8f75c-139">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="8f75c-139">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="8f75c-140">作成</span><span class="sxs-lookup"><span data-stu-id="8f75c-140">Compose</span></span><br><span data-ttu-id="8f75c-141">読み取り</span><span class="sxs-lookup"><span data-stu-id="8f75c-141">Read</span></span> | <span data-ttu-id="8f75c-142">String</span><span class="sxs-lookup"><span data-stu-id="8f75c-142">String</span></span> | [<span data-ttu-id="8f75c-143">1.1</span><span class="sxs-lookup"><span data-stu-id="8f75c-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="8f75c-144">名前空間</span><span class="sxs-lookup"><span data-stu-id="8f75c-144">Namespaces</span></span>

<span data-ttu-id="8f75c-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="8f75c-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="8f75c-146">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="8f75c-146">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="8f75c-147">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="8f75c-147">AsyncResultStatus: String</span></span>

<span data-ttu-id="8f75c-148">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="8f75c-148">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="8f75c-149">型</span><span class="sxs-lookup"><span data-stu-id="8f75c-149">Type</span></span>

*   <span data-ttu-id="8f75c-150">String</span><span class="sxs-lookup"><span data-stu-id="8f75c-150">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8f75c-151">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="8f75c-151">Properties:</span></span>

|<span data-ttu-id="8f75c-152">名前</span><span class="sxs-lookup"><span data-stu-id="8f75c-152">Name</span></span>| <span data-ttu-id="8f75c-153">種類</span><span class="sxs-lookup"><span data-stu-id="8f75c-153">Type</span></span>| <span data-ttu-id="8f75c-154">説明</span><span class="sxs-lookup"><span data-stu-id="8f75c-154">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="8f75c-155">String</span><span class="sxs-lookup"><span data-stu-id="8f75c-155">String</span></span>|<span data-ttu-id="8f75c-156">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="8f75c-156">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="8f75c-157">String</span><span class="sxs-lookup"><span data-stu-id="8f75c-157">String</span></span>|<span data-ttu-id="8f75c-158">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="8f75c-158">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8f75c-159">要件</span><span class="sxs-lookup"><span data-stu-id="8f75c-159">Requirements</span></span>

|<span data-ttu-id="8f75c-160">要件</span><span class="sxs-lookup"><span data-stu-id="8f75c-160">Requirement</span></span>| <span data-ttu-id="8f75c-161">値</span><span class="sxs-lookup"><span data-stu-id="8f75c-161">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f75c-162">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8f75c-162">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f75c-163">1.1</span><span class="sxs-lookup"><span data-stu-id="8f75c-163">1.1</span></span>|
|[<span data-ttu-id="8f75c-164">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8f75c-164">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8f75c-165">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8f75c-165">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="8f75c-166">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="8f75c-166">CoercionType: String</span></span>

<span data-ttu-id="8f75c-167">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="8f75c-167">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8f75c-168">型</span><span class="sxs-lookup"><span data-stu-id="8f75c-168">Type</span></span>

*   <span data-ttu-id="8f75c-169">String</span><span class="sxs-lookup"><span data-stu-id="8f75c-169">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8f75c-170">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="8f75c-170">Properties:</span></span>

|<span data-ttu-id="8f75c-171">名前</span><span class="sxs-lookup"><span data-stu-id="8f75c-171">Name</span></span>| <span data-ttu-id="8f75c-172">種類</span><span class="sxs-lookup"><span data-stu-id="8f75c-172">Type</span></span>| <span data-ttu-id="8f75c-173">説明</span><span class="sxs-lookup"><span data-stu-id="8f75c-173">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="8f75c-174">String</span><span class="sxs-lookup"><span data-stu-id="8f75c-174">String</span></span>|<span data-ttu-id="8f75c-175">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="8f75c-175">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="8f75c-176">String</span><span class="sxs-lookup"><span data-stu-id="8f75c-176">String</span></span>|<span data-ttu-id="8f75c-177">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="8f75c-177">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8f75c-178">要件</span><span class="sxs-lookup"><span data-stu-id="8f75c-178">Requirements</span></span>

|<span data-ttu-id="8f75c-179">要件</span><span class="sxs-lookup"><span data-stu-id="8f75c-179">Requirement</span></span>| <span data-ttu-id="8f75c-180">値</span><span class="sxs-lookup"><span data-stu-id="8f75c-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f75c-181">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8f75c-181">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f75c-182">1.1</span><span class="sxs-lookup"><span data-stu-id="8f75c-182">1.1</span></span>|
|[<span data-ttu-id="8f75c-183">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8f75c-183">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8f75c-184">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8f75c-184">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="8f75c-185">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="8f75c-185">SourceProperty: String</span></span>

<span data-ttu-id="8f75c-186">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="8f75c-186">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8f75c-187">型</span><span class="sxs-lookup"><span data-stu-id="8f75c-187">Type</span></span>

*   <span data-ttu-id="8f75c-188">String</span><span class="sxs-lookup"><span data-stu-id="8f75c-188">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8f75c-189">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="8f75c-189">Properties:</span></span>

|<span data-ttu-id="8f75c-190">名前</span><span class="sxs-lookup"><span data-stu-id="8f75c-190">Name</span></span>| <span data-ttu-id="8f75c-191">種類</span><span class="sxs-lookup"><span data-stu-id="8f75c-191">Type</span></span>| <span data-ttu-id="8f75c-192">説明</span><span class="sxs-lookup"><span data-stu-id="8f75c-192">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="8f75c-193">String</span><span class="sxs-lookup"><span data-stu-id="8f75c-193">String</span></span>|<span data-ttu-id="8f75c-194">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="8f75c-194">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="8f75c-195">String</span><span class="sxs-lookup"><span data-stu-id="8f75c-195">String</span></span>|<span data-ttu-id="8f75c-196">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="8f75c-196">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8f75c-197">要件</span><span class="sxs-lookup"><span data-stu-id="8f75c-197">Requirements</span></span>

|<span data-ttu-id="8f75c-198">要件</span><span class="sxs-lookup"><span data-stu-id="8f75c-198">Requirement</span></span>| <span data-ttu-id="8f75c-199">値</span><span class="sxs-lookup"><span data-stu-id="8f75c-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f75c-200">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8f75c-200">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f75c-201">1.1</span><span class="sxs-lookup"><span data-stu-id="8f75c-201">1.1</span></span>|
|[<span data-ttu-id="8f75c-202">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8f75c-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8f75c-203">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8f75c-203">Compose or Read</span></span>|
