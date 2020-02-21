---
title: Office 名前空間-要件セット1.3
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: b6a8c581670692ed48c4dcc2a7e1f86196b5bce7
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165434"
---
# <a name="office"></a><span data-ttu-id="ab735-102">Office</span><span class="sxs-lookup"><span data-stu-id="ab735-102">Office</span></span>

<span data-ttu-id="ab735-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ab735-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab735-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="ab735-105">Requirements</span></span>

|<span data-ttu-id="ab735-106">要件</span><span class="sxs-lookup"><span data-stu-id="ab735-106">Requirement</span></span>| <span data-ttu-id="ab735-107">値</span><span class="sxs-lookup"><span data-stu-id="ab735-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab735-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ab735-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ab735-109">1.1</span><span class="sxs-lookup"><span data-stu-id="ab735-109">1.1</span></span>|
|[<span data-ttu-id="ab735-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ab735-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ab735-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ab735-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="ab735-112">Properties</span><span class="sxs-lookup"><span data-stu-id="ab735-112">Properties</span></span>

| <span data-ttu-id="ab735-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="ab735-113">Property</span></span> | <span data-ttu-id="ab735-114">モード</span><span class="sxs-lookup"><span data-stu-id="ab735-114">Modes</span></span> | <span data-ttu-id="ab735-115">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="ab735-115">Return type</span></span> | <span data-ttu-id="ab735-116">最小値</span><span class="sxs-lookup"><span data-stu-id="ab735-116">Minimum</span></span><br><span data-ttu-id="ab735-117">要件セット</span><span class="sxs-lookup"><span data-stu-id="ab735-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ab735-118">context</span><span class="sxs-lookup"><span data-stu-id="ab735-118">context</span></span>](office.context.md) | <span data-ttu-id="ab735-119">作成</span><span class="sxs-lookup"><span data-stu-id="ab735-119">Compose</span></span><br><span data-ttu-id="ab735-120">読み取り</span><span class="sxs-lookup"><span data-stu-id="ab735-120">Read</span></span> | [<span data-ttu-id="ab735-121">Context</span><span class="sxs-lookup"><span data-stu-id="ab735-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.3) | [<span data-ttu-id="ab735-122">1.1</span><span class="sxs-lookup"><span data-stu-id="ab735-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="ab735-123">列挙型</span><span class="sxs-lookup"><span data-stu-id="ab735-123">Enumerations</span></span>

| <span data-ttu-id="ab735-124">列挙体</span><span class="sxs-lookup"><span data-stu-id="ab735-124">Enumeration</span></span> | <span data-ttu-id="ab735-125">モード</span><span class="sxs-lookup"><span data-stu-id="ab735-125">Modes</span></span> | <span data-ttu-id="ab735-126">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="ab735-126">Return type</span></span> | <span data-ttu-id="ab735-127">最小値</span><span class="sxs-lookup"><span data-stu-id="ab735-127">Minimum</span></span><br><span data-ttu-id="ab735-128">要件セット</span><span class="sxs-lookup"><span data-stu-id="ab735-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ab735-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="ab735-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="ab735-130">作成</span><span class="sxs-lookup"><span data-stu-id="ab735-130">Compose</span></span><br><span data-ttu-id="ab735-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="ab735-131">Read</span></span> | <span data-ttu-id="ab735-132">文字列</span><span class="sxs-lookup"><span data-stu-id="ab735-132">String</span></span> | [<span data-ttu-id="ab735-133">1.1</span><span class="sxs-lookup"><span data-stu-id="ab735-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ab735-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="ab735-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="ab735-135">作成</span><span class="sxs-lookup"><span data-stu-id="ab735-135">Compose</span></span><br><span data-ttu-id="ab735-136">読み取り</span><span class="sxs-lookup"><span data-stu-id="ab735-136">Read</span></span> | <span data-ttu-id="ab735-137">文字列</span><span class="sxs-lookup"><span data-stu-id="ab735-137">String</span></span> | [<span data-ttu-id="ab735-138">1.1</span><span class="sxs-lookup"><span data-stu-id="ab735-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ab735-139">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="ab735-139">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="ab735-140">作成</span><span class="sxs-lookup"><span data-stu-id="ab735-140">Compose</span></span><br><span data-ttu-id="ab735-141">読み取り</span><span class="sxs-lookup"><span data-stu-id="ab735-141">Read</span></span> | <span data-ttu-id="ab735-142">文字列</span><span class="sxs-lookup"><span data-stu-id="ab735-142">String</span></span> | [<span data-ttu-id="ab735-143">1.1</span><span class="sxs-lookup"><span data-stu-id="ab735-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="ab735-144">名前空間</span><span class="sxs-lookup"><span data-stu-id="ab735-144">Namespaces</span></span>

<span data-ttu-id="ab735-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="ab735-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="ab735-146">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="ab735-146">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="ab735-147">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="ab735-147">AsyncResultStatus: String</span></span>

<span data-ttu-id="ab735-148">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="ab735-148">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="ab735-149">型</span><span class="sxs-lookup"><span data-stu-id="ab735-149">Type</span></span>

*   <span data-ttu-id="ab735-150">String</span><span class="sxs-lookup"><span data-stu-id="ab735-150">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ab735-151">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="ab735-151">Properties:</span></span>

|<span data-ttu-id="ab735-152">名前</span><span class="sxs-lookup"><span data-stu-id="ab735-152">Name</span></span>| <span data-ttu-id="ab735-153">種類</span><span class="sxs-lookup"><span data-stu-id="ab735-153">Type</span></span>| <span data-ttu-id="ab735-154">説明</span><span class="sxs-lookup"><span data-stu-id="ab735-154">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="ab735-155">文字列</span><span class="sxs-lookup"><span data-stu-id="ab735-155">String</span></span>|<span data-ttu-id="ab735-156">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="ab735-156">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="ab735-157">文字列</span><span class="sxs-lookup"><span data-stu-id="ab735-157">String</span></span>|<span data-ttu-id="ab735-158">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="ab735-158">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ab735-159">Requirements</span><span class="sxs-lookup"><span data-stu-id="ab735-159">Requirements</span></span>

|<span data-ttu-id="ab735-160">要件</span><span class="sxs-lookup"><span data-stu-id="ab735-160">Requirement</span></span>| <span data-ttu-id="ab735-161">値</span><span class="sxs-lookup"><span data-stu-id="ab735-161">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab735-162">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ab735-162">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ab735-163">1.1</span><span class="sxs-lookup"><span data-stu-id="ab735-163">1.1</span></span>|
|[<span data-ttu-id="ab735-164">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ab735-164">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ab735-165">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ab735-165">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="ab735-166">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="ab735-166">CoercionType: String</span></span>

<span data-ttu-id="ab735-167">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="ab735-167">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ab735-168">型</span><span class="sxs-lookup"><span data-stu-id="ab735-168">Type</span></span>

*   <span data-ttu-id="ab735-169">String</span><span class="sxs-lookup"><span data-stu-id="ab735-169">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ab735-170">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="ab735-170">Properties:</span></span>

|<span data-ttu-id="ab735-171">名前</span><span class="sxs-lookup"><span data-stu-id="ab735-171">Name</span></span>| <span data-ttu-id="ab735-172">種類</span><span class="sxs-lookup"><span data-stu-id="ab735-172">Type</span></span>| <span data-ttu-id="ab735-173">説明</span><span class="sxs-lookup"><span data-stu-id="ab735-173">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="ab735-174">文字列</span><span class="sxs-lookup"><span data-stu-id="ab735-174">String</span></span>|<span data-ttu-id="ab735-175">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="ab735-175">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="ab735-176">String</span><span class="sxs-lookup"><span data-stu-id="ab735-176">String</span></span>|<span data-ttu-id="ab735-177">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="ab735-177">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ab735-178">Requirements</span><span class="sxs-lookup"><span data-stu-id="ab735-178">Requirements</span></span>

|<span data-ttu-id="ab735-179">要件</span><span class="sxs-lookup"><span data-stu-id="ab735-179">Requirement</span></span>| <span data-ttu-id="ab735-180">値</span><span class="sxs-lookup"><span data-stu-id="ab735-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab735-181">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ab735-181">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ab735-182">1.1</span><span class="sxs-lookup"><span data-stu-id="ab735-182">1.1</span></span>|
|[<span data-ttu-id="ab735-183">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ab735-183">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ab735-184">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ab735-184">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="ab735-185">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="ab735-185">SourceProperty: String</span></span>

<span data-ttu-id="ab735-186">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="ab735-186">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ab735-187">型</span><span class="sxs-lookup"><span data-stu-id="ab735-187">Type</span></span>

*   <span data-ttu-id="ab735-188">String</span><span class="sxs-lookup"><span data-stu-id="ab735-188">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ab735-189">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="ab735-189">Properties:</span></span>

|<span data-ttu-id="ab735-190">名前</span><span class="sxs-lookup"><span data-stu-id="ab735-190">Name</span></span>| <span data-ttu-id="ab735-191">種類</span><span class="sxs-lookup"><span data-stu-id="ab735-191">Type</span></span>| <span data-ttu-id="ab735-192">説明</span><span class="sxs-lookup"><span data-stu-id="ab735-192">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="ab735-193">文字列</span><span class="sxs-lookup"><span data-stu-id="ab735-193">String</span></span>|<span data-ttu-id="ab735-194">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="ab735-194">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="ab735-195">文字列</span><span class="sxs-lookup"><span data-stu-id="ab735-195">String</span></span>|<span data-ttu-id="ab735-196">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="ab735-196">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ab735-197">Requirements</span><span class="sxs-lookup"><span data-stu-id="ab735-197">Requirements</span></span>

|<span data-ttu-id="ab735-198">要件</span><span class="sxs-lookup"><span data-stu-id="ab735-198">Requirement</span></span>| <span data-ttu-id="ab735-199">値</span><span class="sxs-lookup"><span data-stu-id="ab735-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab735-200">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ab735-200">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ab735-201">1.1</span><span class="sxs-lookup"><span data-stu-id="ab735-201">1.1</span></span>|
|[<span data-ttu-id="ab735-202">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ab735-202">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ab735-203">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ab735-203">Compose or Read</span></span>|
