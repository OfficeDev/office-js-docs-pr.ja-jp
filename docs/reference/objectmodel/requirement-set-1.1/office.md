---
title: Office 名前空間 - 要件セット 1.1
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 68363f101b4c818853cc118e39d05784c56ef3ad
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165476"
---
# <a name="office"></a><span data-ttu-id="dd0a1-102">Office</span><span class="sxs-lookup"><span data-stu-id="dd0a1-102">Office</span></span>

<span data-ttu-id="dd0a1-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="dd0a1-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0a1-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="dd0a1-105">Requirements</span></span>

|<span data-ttu-id="dd0a1-106">要件</span><span class="sxs-lookup"><span data-stu-id="dd0a1-106">Requirement</span></span>| <span data-ttu-id="dd0a1-107">値</span><span class="sxs-lookup"><span data-stu-id="dd0a1-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0a1-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0a1-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dd0a1-109">1.1</span><span class="sxs-lookup"><span data-stu-id="dd0a1-109">1.1</span></span>|
|[<span data-ttu-id="dd0a1-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0a1-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dd0a1-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dd0a1-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="dd0a1-112">Properties</span><span class="sxs-lookup"><span data-stu-id="dd0a1-112">Properties</span></span>

| <span data-ttu-id="dd0a1-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="dd0a1-113">Property</span></span> | <span data-ttu-id="dd0a1-114">モード</span><span class="sxs-lookup"><span data-stu-id="dd0a1-114">Modes</span></span> | <span data-ttu-id="dd0a1-115">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="dd0a1-115">Return type</span></span> | <span data-ttu-id="dd0a1-116">最小値</span><span class="sxs-lookup"><span data-stu-id="dd0a1-116">Minimum</span></span><br><span data-ttu-id="dd0a1-117">要件セット</span><span class="sxs-lookup"><span data-stu-id="dd0a1-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="dd0a1-118">context</span><span class="sxs-lookup"><span data-stu-id="dd0a1-118">context</span></span>](office.context.md) | <span data-ttu-id="dd0a1-119">作成</span><span class="sxs-lookup"><span data-stu-id="dd0a1-119">Compose</span></span><br><span data-ttu-id="dd0a1-120">読み取り</span><span class="sxs-lookup"><span data-stu-id="dd0a1-120">Read</span></span> | [<span data-ttu-id="dd0a1-121">Context</span><span class="sxs-lookup"><span data-stu-id="dd0a1-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.1) | [<span data-ttu-id="dd0a1-122">1.1</span><span class="sxs-lookup"><span data-stu-id="dd0a1-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="dd0a1-123">列挙型</span><span class="sxs-lookup"><span data-stu-id="dd0a1-123">Enumerations</span></span>

| <span data-ttu-id="dd0a1-124">列挙体</span><span class="sxs-lookup"><span data-stu-id="dd0a1-124">Enumeration</span></span> | <span data-ttu-id="dd0a1-125">モード</span><span class="sxs-lookup"><span data-stu-id="dd0a1-125">Modes</span></span> | <span data-ttu-id="dd0a1-126">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="dd0a1-126">Return type</span></span> | <span data-ttu-id="dd0a1-127">最小値</span><span class="sxs-lookup"><span data-stu-id="dd0a1-127">Minimum</span></span><br><span data-ttu-id="dd0a1-128">要件セット</span><span class="sxs-lookup"><span data-stu-id="dd0a1-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="dd0a1-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="dd0a1-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="dd0a1-130">作成</span><span class="sxs-lookup"><span data-stu-id="dd0a1-130">Compose</span></span><br><span data-ttu-id="dd0a1-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="dd0a1-131">Read</span></span> | <span data-ttu-id="dd0a1-132">文字列</span><span class="sxs-lookup"><span data-stu-id="dd0a1-132">String</span></span> | [<span data-ttu-id="dd0a1-133">1.1</span><span class="sxs-lookup"><span data-stu-id="dd0a1-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dd0a1-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="dd0a1-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="dd0a1-135">作成</span><span class="sxs-lookup"><span data-stu-id="dd0a1-135">Compose</span></span><br><span data-ttu-id="dd0a1-136">読み取り</span><span class="sxs-lookup"><span data-stu-id="dd0a1-136">Read</span></span> | <span data-ttu-id="dd0a1-137">文字列</span><span class="sxs-lookup"><span data-stu-id="dd0a1-137">String</span></span> | [<span data-ttu-id="dd0a1-138">1.1</span><span class="sxs-lookup"><span data-stu-id="dd0a1-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dd0a1-139">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="dd0a1-139">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="dd0a1-140">作成</span><span class="sxs-lookup"><span data-stu-id="dd0a1-140">Compose</span></span><br><span data-ttu-id="dd0a1-141">読み取り</span><span class="sxs-lookup"><span data-stu-id="dd0a1-141">Read</span></span> | <span data-ttu-id="dd0a1-142">文字列</span><span class="sxs-lookup"><span data-stu-id="dd0a1-142">String</span></span> | [<span data-ttu-id="dd0a1-143">1.1</span><span class="sxs-lookup"><span data-stu-id="dd0a1-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="dd0a1-144">名前空間</span><span class="sxs-lookup"><span data-stu-id="dd0a1-144">Namespaces</span></span>

<span data-ttu-id="dd0a1-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="dd0a1-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="dd0a1-146">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="dd0a1-146">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="dd0a1-147">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="dd0a1-147">AsyncResultStatus: String</span></span>

<span data-ttu-id="dd0a1-148">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="dd0a1-148">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="dd0a1-149">型</span><span class="sxs-lookup"><span data-stu-id="dd0a1-149">Type</span></span>

*   <span data-ttu-id="dd0a1-150">String</span><span class="sxs-lookup"><span data-stu-id="dd0a1-150">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dd0a1-151">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="dd0a1-151">Properties:</span></span>

|<span data-ttu-id="dd0a1-152">名前</span><span class="sxs-lookup"><span data-stu-id="dd0a1-152">Name</span></span>| <span data-ttu-id="dd0a1-153">種類</span><span class="sxs-lookup"><span data-stu-id="dd0a1-153">Type</span></span>| <span data-ttu-id="dd0a1-154">説明</span><span class="sxs-lookup"><span data-stu-id="dd0a1-154">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="dd0a1-155">文字列</span><span class="sxs-lookup"><span data-stu-id="dd0a1-155">String</span></span>|<span data-ttu-id="dd0a1-156">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="dd0a1-156">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="dd0a1-157">文字列</span><span class="sxs-lookup"><span data-stu-id="dd0a1-157">String</span></span>|<span data-ttu-id="dd0a1-158">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="dd0a1-158">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd0a1-159">Requirements</span><span class="sxs-lookup"><span data-stu-id="dd0a1-159">Requirements</span></span>

|<span data-ttu-id="dd0a1-160">要件</span><span class="sxs-lookup"><span data-stu-id="dd0a1-160">Requirement</span></span>| <span data-ttu-id="dd0a1-161">値</span><span class="sxs-lookup"><span data-stu-id="dd0a1-161">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0a1-162">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0a1-162">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dd0a1-163">1.1</span><span class="sxs-lookup"><span data-stu-id="dd0a1-163">1.1</span></span>|
|[<span data-ttu-id="dd0a1-164">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0a1-164">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dd0a1-165">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dd0a1-165">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="dd0a1-166">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="dd0a1-166">CoercionType: String</span></span>

<span data-ttu-id="dd0a1-167">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="dd0a1-167">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="dd0a1-168">型</span><span class="sxs-lookup"><span data-stu-id="dd0a1-168">Type</span></span>

*   <span data-ttu-id="dd0a1-169">String</span><span class="sxs-lookup"><span data-stu-id="dd0a1-169">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dd0a1-170">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="dd0a1-170">Properties:</span></span>

|<span data-ttu-id="dd0a1-171">名前</span><span class="sxs-lookup"><span data-stu-id="dd0a1-171">Name</span></span>| <span data-ttu-id="dd0a1-172">種類</span><span class="sxs-lookup"><span data-stu-id="dd0a1-172">Type</span></span>| <span data-ttu-id="dd0a1-173">説明</span><span class="sxs-lookup"><span data-stu-id="dd0a1-173">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="dd0a1-174">文字列</span><span class="sxs-lookup"><span data-stu-id="dd0a1-174">String</span></span>|<span data-ttu-id="dd0a1-175">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="dd0a1-175">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="dd0a1-176">String</span><span class="sxs-lookup"><span data-stu-id="dd0a1-176">String</span></span>|<span data-ttu-id="dd0a1-177">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="dd0a1-177">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd0a1-178">Requirements</span><span class="sxs-lookup"><span data-stu-id="dd0a1-178">Requirements</span></span>

|<span data-ttu-id="dd0a1-179">要件</span><span class="sxs-lookup"><span data-stu-id="dd0a1-179">Requirement</span></span>| <span data-ttu-id="dd0a1-180">値</span><span class="sxs-lookup"><span data-stu-id="dd0a1-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0a1-181">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0a1-181">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dd0a1-182">1.1</span><span class="sxs-lookup"><span data-stu-id="dd0a1-182">1.1</span></span>|
|[<span data-ttu-id="dd0a1-183">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0a1-183">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dd0a1-184">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dd0a1-184">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="dd0a1-185">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="dd0a1-185">SourceProperty: String</span></span>

<span data-ttu-id="dd0a1-186">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="dd0a1-186">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="dd0a1-187">型</span><span class="sxs-lookup"><span data-stu-id="dd0a1-187">Type</span></span>

*   <span data-ttu-id="dd0a1-188">String</span><span class="sxs-lookup"><span data-stu-id="dd0a1-188">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dd0a1-189">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="dd0a1-189">Properties:</span></span>

|<span data-ttu-id="dd0a1-190">名前</span><span class="sxs-lookup"><span data-stu-id="dd0a1-190">Name</span></span>| <span data-ttu-id="dd0a1-191">種類</span><span class="sxs-lookup"><span data-stu-id="dd0a1-191">Type</span></span>| <span data-ttu-id="dd0a1-192">説明</span><span class="sxs-lookup"><span data-stu-id="dd0a1-192">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="dd0a1-193">文字列</span><span class="sxs-lookup"><span data-stu-id="dd0a1-193">String</span></span>|<span data-ttu-id="dd0a1-194">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="dd0a1-194">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="dd0a1-195">文字列</span><span class="sxs-lookup"><span data-stu-id="dd0a1-195">String</span></span>|<span data-ttu-id="dd0a1-196">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="dd0a1-196">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd0a1-197">Requirements</span><span class="sxs-lookup"><span data-stu-id="dd0a1-197">Requirements</span></span>

|<span data-ttu-id="dd0a1-198">要件</span><span class="sxs-lookup"><span data-stu-id="dd0a1-198">Requirement</span></span>| <span data-ttu-id="dd0a1-199">値</span><span class="sxs-lookup"><span data-stu-id="dd0a1-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0a1-200">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0a1-200">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dd0a1-201">1.1</span><span class="sxs-lookup"><span data-stu-id="dd0a1-201">1.1</span></span>|
|[<span data-ttu-id="dd0a1-202">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0a1-202">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dd0a1-203">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dd0a1-203">Compose or Read</span></span>|