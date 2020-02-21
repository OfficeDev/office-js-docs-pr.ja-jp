---
title: Office 名前空間-要件セット1.2
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 0f955ed8279655b4ac92dc04871a1227b045f6ea
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165441"
---
# <a name="office"></a><span data-ttu-id="f0dca-102">Office</span><span class="sxs-lookup"><span data-stu-id="f0dca-102">Office</span></span>

<span data-ttu-id="f0dca-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f0dca-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0dca-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0dca-105">Requirements</span></span>

|<span data-ttu-id="f0dca-106">要件</span><span class="sxs-lookup"><span data-stu-id="f0dca-106">Requirement</span></span>| <span data-ttu-id="f0dca-107">値</span><span class="sxs-lookup"><span data-stu-id="f0dca-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0dca-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0dca-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f0dca-109">1.1</span><span class="sxs-lookup"><span data-stu-id="f0dca-109">1.1</span></span>|
|[<span data-ttu-id="f0dca-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0dca-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f0dca-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0dca-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="f0dca-112">Properties</span><span class="sxs-lookup"><span data-stu-id="f0dca-112">Properties</span></span>

| <span data-ttu-id="f0dca-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="f0dca-113">Property</span></span> | <span data-ttu-id="f0dca-114">モード</span><span class="sxs-lookup"><span data-stu-id="f0dca-114">Modes</span></span> | <span data-ttu-id="f0dca-115">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="f0dca-115">Return type</span></span> | <span data-ttu-id="f0dca-116">最小値</span><span class="sxs-lookup"><span data-stu-id="f0dca-116">Minimum</span></span><br><span data-ttu-id="f0dca-117">要件セット</span><span class="sxs-lookup"><span data-stu-id="f0dca-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f0dca-118">context</span><span class="sxs-lookup"><span data-stu-id="f0dca-118">context</span></span>](office.context.md) | <span data-ttu-id="f0dca-119">作成</span><span class="sxs-lookup"><span data-stu-id="f0dca-119">Compose</span></span><br><span data-ttu-id="f0dca-120">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0dca-120">Read</span></span> | [<span data-ttu-id="f0dca-121">Context</span><span class="sxs-lookup"><span data-stu-id="f0dca-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.2) | [<span data-ttu-id="f0dca-122">1.1</span><span class="sxs-lookup"><span data-stu-id="f0dca-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="f0dca-123">列挙型</span><span class="sxs-lookup"><span data-stu-id="f0dca-123">Enumerations</span></span>

| <span data-ttu-id="f0dca-124">列挙体</span><span class="sxs-lookup"><span data-stu-id="f0dca-124">Enumeration</span></span> | <span data-ttu-id="f0dca-125">モード</span><span class="sxs-lookup"><span data-stu-id="f0dca-125">Modes</span></span> | <span data-ttu-id="f0dca-126">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="f0dca-126">Return type</span></span> | <span data-ttu-id="f0dca-127">最小値</span><span class="sxs-lookup"><span data-stu-id="f0dca-127">Minimum</span></span><br><span data-ttu-id="f0dca-128">要件セット</span><span class="sxs-lookup"><span data-stu-id="f0dca-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f0dca-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="f0dca-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="f0dca-130">作成</span><span class="sxs-lookup"><span data-stu-id="f0dca-130">Compose</span></span><br><span data-ttu-id="f0dca-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0dca-131">Read</span></span> | <span data-ttu-id="f0dca-132">文字列</span><span class="sxs-lookup"><span data-stu-id="f0dca-132">String</span></span> | [<span data-ttu-id="f0dca-133">1.1</span><span class="sxs-lookup"><span data-stu-id="f0dca-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f0dca-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="f0dca-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="f0dca-135">作成</span><span class="sxs-lookup"><span data-stu-id="f0dca-135">Compose</span></span><br><span data-ttu-id="f0dca-136">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0dca-136">Read</span></span> | <span data-ttu-id="f0dca-137">文字列</span><span class="sxs-lookup"><span data-stu-id="f0dca-137">String</span></span> | [<span data-ttu-id="f0dca-138">1.1</span><span class="sxs-lookup"><span data-stu-id="f0dca-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f0dca-139">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="f0dca-139">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="f0dca-140">作成</span><span class="sxs-lookup"><span data-stu-id="f0dca-140">Compose</span></span><br><span data-ttu-id="f0dca-141">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0dca-141">Read</span></span> | <span data-ttu-id="f0dca-142">文字列</span><span class="sxs-lookup"><span data-stu-id="f0dca-142">String</span></span> | [<span data-ttu-id="f0dca-143">1.1</span><span class="sxs-lookup"><span data-stu-id="f0dca-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="f0dca-144">名前空間</span><span class="sxs-lookup"><span data-stu-id="f0dca-144">Namespaces</span></span>

<span data-ttu-id="f0dca-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="f0dca-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="f0dca-146">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="f0dca-146">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="f0dca-147">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="f0dca-147">AsyncResultStatus: String</span></span>

<span data-ttu-id="f0dca-148">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="f0dca-148">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f0dca-149">型</span><span class="sxs-lookup"><span data-stu-id="f0dca-149">Type</span></span>

*   <span data-ttu-id="f0dca-150">String</span><span class="sxs-lookup"><span data-stu-id="f0dca-150">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f0dca-151">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="f0dca-151">Properties:</span></span>

|<span data-ttu-id="f0dca-152">名前</span><span class="sxs-lookup"><span data-stu-id="f0dca-152">Name</span></span>| <span data-ttu-id="f0dca-153">種類</span><span class="sxs-lookup"><span data-stu-id="f0dca-153">Type</span></span>| <span data-ttu-id="f0dca-154">説明</span><span class="sxs-lookup"><span data-stu-id="f0dca-154">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f0dca-155">文字列</span><span class="sxs-lookup"><span data-stu-id="f0dca-155">String</span></span>|<span data-ttu-id="f0dca-156">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="f0dca-156">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f0dca-157">文字列</span><span class="sxs-lookup"><span data-stu-id="f0dca-157">String</span></span>|<span data-ttu-id="f0dca-158">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="f0dca-158">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f0dca-159">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0dca-159">Requirements</span></span>

|<span data-ttu-id="f0dca-160">要件</span><span class="sxs-lookup"><span data-stu-id="f0dca-160">Requirement</span></span>| <span data-ttu-id="f0dca-161">値</span><span class="sxs-lookup"><span data-stu-id="f0dca-161">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0dca-162">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0dca-162">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f0dca-163">1.1</span><span class="sxs-lookup"><span data-stu-id="f0dca-163">1.1</span></span>|
|[<span data-ttu-id="f0dca-164">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0dca-164">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f0dca-165">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0dca-165">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="f0dca-166">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="f0dca-166">CoercionType: String</span></span>

<span data-ttu-id="f0dca-167">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="f0dca-167">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f0dca-168">型</span><span class="sxs-lookup"><span data-stu-id="f0dca-168">Type</span></span>

*   <span data-ttu-id="f0dca-169">String</span><span class="sxs-lookup"><span data-stu-id="f0dca-169">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f0dca-170">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="f0dca-170">Properties:</span></span>

|<span data-ttu-id="f0dca-171">名前</span><span class="sxs-lookup"><span data-stu-id="f0dca-171">Name</span></span>| <span data-ttu-id="f0dca-172">種類</span><span class="sxs-lookup"><span data-stu-id="f0dca-172">Type</span></span>| <span data-ttu-id="f0dca-173">説明</span><span class="sxs-lookup"><span data-stu-id="f0dca-173">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f0dca-174">文字列</span><span class="sxs-lookup"><span data-stu-id="f0dca-174">String</span></span>|<span data-ttu-id="f0dca-175">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="f0dca-175">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f0dca-176">String</span><span class="sxs-lookup"><span data-stu-id="f0dca-176">String</span></span>|<span data-ttu-id="f0dca-177">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="f0dca-177">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f0dca-178">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0dca-178">Requirements</span></span>

|<span data-ttu-id="f0dca-179">要件</span><span class="sxs-lookup"><span data-stu-id="f0dca-179">Requirement</span></span>| <span data-ttu-id="f0dca-180">値</span><span class="sxs-lookup"><span data-stu-id="f0dca-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0dca-181">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0dca-181">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f0dca-182">1.1</span><span class="sxs-lookup"><span data-stu-id="f0dca-182">1.1</span></span>|
|[<span data-ttu-id="f0dca-183">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0dca-183">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f0dca-184">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0dca-184">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="f0dca-185">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="f0dca-185">SourceProperty: String</span></span>

<span data-ttu-id="f0dca-186">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="f0dca-186">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f0dca-187">型</span><span class="sxs-lookup"><span data-stu-id="f0dca-187">Type</span></span>

*   <span data-ttu-id="f0dca-188">String</span><span class="sxs-lookup"><span data-stu-id="f0dca-188">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f0dca-189">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="f0dca-189">Properties:</span></span>

|<span data-ttu-id="f0dca-190">名前</span><span class="sxs-lookup"><span data-stu-id="f0dca-190">Name</span></span>| <span data-ttu-id="f0dca-191">種類</span><span class="sxs-lookup"><span data-stu-id="f0dca-191">Type</span></span>| <span data-ttu-id="f0dca-192">説明</span><span class="sxs-lookup"><span data-stu-id="f0dca-192">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f0dca-193">文字列</span><span class="sxs-lookup"><span data-stu-id="f0dca-193">String</span></span>|<span data-ttu-id="f0dca-194">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="f0dca-194">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f0dca-195">文字列</span><span class="sxs-lookup"><span data-stu-id="f0dca-195">String</span></span>|<span data-ttu-id="f0dca-196">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="f0dca-196">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f0dca-197">Requirements</span><span class="sxs-lookup"><span data-stu-id="f0dca-197">Requirements</span></span>

|<span data-ttu-id="f0dca-198">要件</span><span class="sxs-lookup"><span data-stu-id="f0dca-198">Requirement</span></span>| <span data-ttu-id="f0dca-199">値</span><span class="sxs-lookup"><span data-stu-id="f0dca-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0dca-200">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0dca-200">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f0dca-201">1.1</span><span class="sxs-lookup"><span data-stu-id="f0dca-201">1.1</span></span>|
|[<span data-ttu-id="f0dca-202">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0dca-202">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f0dca-203">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0dca-203">Compose or Read</span></span>|
