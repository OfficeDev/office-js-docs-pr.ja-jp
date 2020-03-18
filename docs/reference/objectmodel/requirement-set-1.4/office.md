---
title: Office 名前空間-要件セット1.4
description: Outlook アドイン API の最上位レベルの名前空間のオブジェクトモデル (Mailbox API 1.4 バージョン)。
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: e5a5c6de5bb87cb32968d9d9d80c621f0acc238d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720058"
---
# <a name="office"></a><span data-ttu-id="41619-103">Office</span><span class="sxs-lookup"><span data-stu-id="41619-103">Office</span></span>

<span data-ttu-id="41619-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="41619-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="41619-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="41619-106">Requirements</span></span>

|<span data-ttu-id="41619-107">要件</span><span class="sxs-lookup"><span data-stu-id="41619-107">Requirement</span></span>| <span data-ttu-id="41619-108">値</span><span class="sxs-lookup"><span data-stu-id="41619-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="41619-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="41619-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="41619-110">1.1</span><span class="sxs-lookup"><span data-stu-id="41619-110">1.1</span></span>|
|[<span data-ttu-id="41619-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="41619-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="41619-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="41619-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="41619-113">Properties</span><span class="sxs-lookup"><span data-stu-id="41619-113">Properties</span></span>

| <span data-ttu-id="41619-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="41619-114">Property</span></span> | <span data-ttu-id="41619-115">モード</span><span class="sxs-lookup"><span data-stu-id="41619-115">Modes</span></span> | <span data-ttu-id="41619-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="41619-116">Return type</span></span> | <span data-ttu-id="41619-117">最小値</span><span class="sxs-lookup"><span data-stu-id="41619-117">Minimum</span></span><br><span data-ttu-id="41619-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="41619-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="41619-119">context</span><span class="sxs-lookup"><span data-stu-id="41619-119">context</span></span>](office.context.md) | <span data-ttu-id="41619-120">作成</span><span class="sxs-lookup"><span data-stu-id="41619-120">Compose</span></span><br><span data-ttu-id="41619-121">読み取り</span><span class="sxs-lookup"><span data-stu-id="41619-121">Read</span></span> | [<span data-ttu-id="41619-122">Context</span><span class="sxs-lookup"><span data-stu-id="41619-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.4) | [<span data-ttu-id="41619-123">1.1</span><span class="sxs-lookup"><span data-stu-id="41619-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="41619-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="41619-124">Enumerations</span></span>

| <span data-ttu-id="41619-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="41619-125">Enumeration</span></span> | <span data-ttu-id="41619-126">モード</span><span class="sxs-lookup"><span data-stu-id="41619-126">Modes</span></span> | <span data-ttu-id="41619-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="41619-127">Return type</span></span> | <span data-ttu-id="41619-128">最小値</span><span class="sxs-lookup"><span data-stu-id="41619-128">Minimum</span></span><br><span data-ttu-id="41619-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="41619-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="41619-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="41619-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="41619-131">作成</span><span class="sxs-lookup"><span data-stu-id="41619-131">Compose</span></span><br><span data-ttu-id="41619-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="41619-132">Read</span></span> | <span data-ttu-id="41619-133">文字列</span><span class="sxs-lookup"><span data-stu-id="41619-133">String</span></span> | [<span data-ttu-id="41619-134">1.1</span><span class="sxs-lookup"><span data-stu-id="41619-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="41619-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="41619-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="41619-136">作成</span><span class="sxs-lookup"><span data-stu-id="41619-136">Compose</span></span><br><span data-ttu-id="41619-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="41619-137">Read</span></span> | <span data-ttu-id="41619-138">文字列</span><span class="sxs-lookup"><span data-stu-id="41619-138">String</span></span> | [<span data-ttu-id="41619-139">1.1</span><span class="sxs-lookup"><span data-stu-id="41619-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="41619-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="41619-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="41619-141">作成</span><span class="sxs-lookup"><span data-stu-id="41619-141">Compose</span></span><br><span data-ttu-id="41619-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="41619-142">Read</span></span> | <span data-ttu-id="41619-143">文字列</span><span class="sxs-lookup"><span data-stu-id="41619-143">String</span></span> | [<span data-ttu-id="41619-144">1.1</span><span class="sxs-lookup"><span data-stu-id="41619-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="41619-145">名前空間</span><span class="sxs-lookup"><span data-stu-id="41619-145">Namespaces</span></span>

<span data-ttu-id="41619-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="41619-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="41619-147">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="41619-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="41619-148">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="41619-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="41619-149">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="41619-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="41619-150">型</span><span class="sxs-lookup"><span data-stu-id="41619-150">Type</span></span>

*   <span data-ttu-id="41619-151">String</span><span class="sxs-lookup"><span data-stu-id="41619-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="41619-152">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="41619-152">Properties:</span></span>

|<span data-ttu-id="41619-153">名前</span><span class="sxs-lookup"><span data-stu-id="41619-153">Name</span></span>| <span data-ttu-id="41619-154">種類</span><span class="sxs-lookup"><span data-stu-id="41619-154">Type</span></span>| <span data-ttu-id="41619-155">説明</span><span class="sxs-lookup"><span data-stu-id="41619-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="41619-156">文字列</span><span class="sxs-lookup"><span data-stu-id="41619-156">String</span></span>|<span data-ttu-id="41619-157">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="41619-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="41619-158">文字列</span><span class="sxs-lookup"><span data-stu-id="41619-158">String</span></span>|<span data-ttu-id="41619-159">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="41619-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="41619-160">Requirements</span><span class="sxs-lookup"><span data-stu-id="41619-160">Requirements</span></span>

|<span data-ttu-id="41619-161">要件</span><span class="sxs-lookup"><span data-stu-id="41619-161">Requirement</span></span>| <span data-ttu-id="41619-162">値</span><span class="sxs-lookup"><span data-stu-id="41619-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="41619-163">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="41619-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="41619-164">1.1</span><span class="sxs-lookup"><span data-stu-id="41619-164">1.1</span></span>|
|[<span data-ttu-id="41619-165">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="41619-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="41619-166">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="41619-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="41619-167">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="41619-167">CoercionType: String</span></span>

<span data-ttu-id="41619-168">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="41619-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="41619-169">型</span><span class="sxs-lookup"><span data-stu-id="41619-169">Type</span></span>

*   <span data-ttu-id="41619-170">String</span><span class="sxs-lookup"><span data-stu-id="41619-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="41619-171">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="41619-171">Properties:</span></span>

|<span data-ttu-id="41619-172">名前</span><span class="sxs-lookup"><span data-stu-id="41619-172">Name</span></span>| <span data-ttu-id="41619-173">種類</span><span class="sxs-lookup"><span data-stu-id="41619-173">Type</span></span>| <span data-ttu-id="41619-174">説明</span><span class="sxs-lookup"><span data-stu-id="41619-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="41619-175">文字列</span><span class="sxs-lookup"><span data-stu-id="41619-175">String</span></span>|<span data-ttu-id="41619-176">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="41619-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="41619-177">String</span><span class="sxs-lookup"><span data-stu-id="41619-177">String</span></span>|<span data-ttu-id="41619-178">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="41619-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="41619-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="41619-179">Requirements</span></span>

|<span data-ttu-id="41619-180">要件</span><span class="sxs-lookup"><span data-stu-id="41619-180">Requirement</span></span>| <span data-ttu-id="41619-181">値</span><span class="sxs-lookup"><span data-stu-id="41619-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="41619-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="41619-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="41619-183">1.1</span><span class="sxs-lookup"><span data-stu-id="41619-183">1.1</span></span>|
|[<span data-ttu-id="41619-184">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="41619-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="41619-185">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="41619-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="41619-186">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="41619-186">SourceProperty: String</span></span>

<span data-ttu-id="41619-187">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="41619-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="41619-188">型</span><span class="sxs-lookup"><span data-stu-id="41619-188">Type</span></span>

*   <span data-ttu-id="41619-189">String</span><span class="sxs-lookup"><span data-stu-id="41619-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="41619-190">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="41619-190">Properties:</span></span>

|<span data-ttu-id="41619-191">名前</span><span class="sxs-lookup"><span data-stu-id="41619-191">Name</span></span>| <span data-ttu-id="41619-192">種類</span><span class="sxs-lookup"><span data-stu-id="41619-192">Type</span></span>| <span data-ttu-id="41619-193">説明</span><span class="sxs-lookup"><span data-stu-id="41619-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="41619-194">文字列</span><span class="sxs-lookup"><span data-stu-id="41619-194">String</span></span>|<span data-ttu-id="41619-195">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="41619-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="41619-196">文字列</span><span class="sxs-lookup"><span data-stu-id="41619-196">String</span></span>|<span data-ttu-id="41619-197">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="41619-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="41619-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="41619-198">Requirements</span></span>

|<span data-ttu-id="41619-199">要件</span><span class="sxs-lookup"><span data-stu-id="41619-199">Requirement</span></span>| <span data-ttu-id="41619-200">値</span><span class="sxs-lookup"><span data-stu-id="41619-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="41619-201">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="41619-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="41619-202">1.1</span><span class="sxs-lookup"><span data-stu-id="41619-202">1.1</span></span>|
|[<span data-ttu-id="41619-203">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="41619-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="41619-204">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="41619-204">Compose or Read</span></span>|
