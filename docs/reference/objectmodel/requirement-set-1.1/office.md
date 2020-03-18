---
title: Office 名前空間 - 要件セット 1.1
description: Outlook アドイン API の最上位レベルの名前空間のオブジェクトモデル (Mailbox API 1.1 バージョン)。
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: e881f0f9eac054f2b95436504da24cc7d4dec86d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720191"
---
# <a name="office"></a><span data-ttu-id="8243f-103">Office</span><span class="sxs-lookup"><span data-stu-id="8243f-103">Office</span></span>

<span data-ttu-id="8243f-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8243f-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8243f-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="8243f-106">Requirements</span></span>

|<span data-ttu-id="8243f-107">要件</span><span class="sxs-lookup"><span data-stu-id="8243f-107">Requirement</span></span>| <span data-ttu-id="8243f-108">値</span><span class="sxs-lookup"><span data-stu-id="8243f-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="8243f-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8243f-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8243f-110">1.1</span><span class="sxs-lookup"><span data-stu-id="8243f-110">1.1</span></span>|
|[<span data-ttu-id="8243f-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8243f-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8243f-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8243f-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="8243f-113">Properties</span><span class="sxs-lookup"><span data-stu-id="8243f-113">Properties</span></span>

| <span data-ttu-id="8243f-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="8243f-114">Property</span></span> | <span data-ttu-id="8243f-115">モード</span><span class="sxs-lookup"><span data-stu-id="8243f-115">Modes</span></span> | <span data-ttu-id="8243f-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="8243f-116">Return type</span></span> | <span data-ttu-id="8243f-117">最小値</span><span class="sxs-lookup"><span data-stu-id="8243f-117">Minimum</span></span><br><span data-ttu-id="8243f-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="8243f-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8243f-119">context</span><span class="sxs-lookup"><span data-stu-id="8243f-119">context</span></span>](office.context.md) | <span data-ttu-id="8243f-120">作成</span><span class="sxs-lookup"><span data-stu-id="8243f-120">Compose</span></span><br><span data-ttu-id="8243f-121">読み取り</span><span class="sxs-lookup"><span data-stu-id="8243f-121">Read</span></span> | [<span data-ttu-id="8243f-122">Context</span><span class="sxs-lookup"><span data-stu-id="8243f-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.1) | [<span data-ttu-id="8243f-123">1.1</span><span class="sxs-lookup"><span data-stu-id="8243f-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="8243f-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="8243f-124">Enumerations</span></span>

| <span data-ttu-id="8243f-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="8243f-125">Enumeration</span></span> | <span data-ttu-id="8243f-126">モード</span><span class="sxs-lookup"><span data-stu-id="8243f-126">Modes</span></span> | <span data-ttu-id="8243f-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="8243f-127">Return type</span></span> | <span data-ttu-id="8243f-128">最小値</span><span class="sxs-lookup"><span data-stu-id="8243f-128">Minimum</span></span><br><span data-ttu-id="8243f-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="8243f-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8243f-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="8243f-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="8243f-131">作成</span><span class="sxs-lookup"><span data-stu-id="8243f-131">Compose</span></span><br><span data-ttu-id="8243f-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="8243f-132">Read</span></span> | <span data-ttu-id="8243f-133">文字列</span><span class="sxs-lookup"><span data-stu-id="8243f-133">String</span></span> | [<span data-ttu-id="8243f-134">1.1</span><span class="sxs-lookup"><span data-stu-id="8243f-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8243f-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="8243f-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="8243f-136">作成</span><span class="sxs-lookup"><span data-stu-id="8243f-136">Compose</span></span><br><span data-ttu-id="8243f-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="8243f-137">Read</span></span> | <span data-ttu-id="8243f-138">文字列</span><span class="sxs-lookup"><span data-stu-id="8243f-138">String</span></span> | [<span data-ttu-id="8243f-139">1.1</span><span class="sxs-lookup"><span data-stu-id="8243f-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8243f-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="8243f-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="8243f-141">作成</span><span class="sxs-lookup"><span data-stu-id="8243f-141">Compose</span></span><br><span data-ttu-id="8243f-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="8243f-142">Read</span></span> | <span data-ttu-id="8243f-143">文字列</span><span class="sxs-lookup"><span data-stu-id="8243f-143">String</span></span> | [<span data-ttu-id="8243f-144">1.1</span><span class="sxs-lookup"><span data-stu-id="8243f-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="8243f-145">名前空間</span><span class="sxs-lookup"><span data-stu-id="8243f-145">Namespaces</span></span>

<span data-ttu-id="8243f-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="8243f-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="8243f-147">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="8243f-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="8243f-148">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="8243f-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="8243f-149">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="8243f-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="8243f-150">型</span><span class="sxs-lookup"><span data-stu-id="8243f-150">Type</span></span>

*   <span data-ttu-id="8243f-151">String</span><span class="sxs-lookup"><span data-stu-id="8243f-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8243f-152">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="8243f-152">Properties:</span></span>

|<span data-ttu-id="8243f-153">名前</span><span class="sxs-lookup"><span data-stu-id="8243f-153">Name</span></span>| <span data-ttu-id="8243f-154">種類</span><span class="sxs-lookup"><span data-stu-id="8243f-154">Type</span></span>| <span data-ttu-id="8243f-155">説明</span><span class="sxs-lookup"><span data-stu-id="8243f-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="8243f-156">文字列</span><span class="sxs-lookup"><span data-stu-id="8243f-156">String</span></span>|<span data-ttu-id="8243f-157">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="8243f-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="8243f-158">文字列</span><span class="sxs-lookup"><span data-stu-id="8243f-158">String</span></span>|<span data-ttu-id="8243f-159">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="8243f-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8243f-160">Requirements</span><span class="sxs-lookup"><span data-stu-id="8243f-160">Requirements</span></span>

|<span data-ttu-id="8243f-161">要件</span><span class="sxs-lookup"><span data-stu-id="8243f-161">Requirement</span></span>| <span data-ttu-id="8243f-162">値</span><span class="sxs-lookup"><span data-stu-id="8243f-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="8243f-163">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8243f-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8243f-164">1.1</span><span class="sxs-lookup"><span data-stu-id="8243f-164">1.1</span></span>|
|[<span data-ttu-id="8243f-165">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8243f-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8243f-166">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8243f-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="8243f-167">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="8243f-167">CoercionType: String</span></span>

<span data-ttu-id="8243f-168">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="8243f-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8243f-169">型</span><span class="sxs-lookup"><span data-stu-id="8243f-169">Type</span></span>

*   <span data-ttu-id="8243f-170">String</span><span class="sxs-lookup"><span data-stu-id="8243f-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8243f-171">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="8243f-171">Properties:</span></span>

|<span data-ttu-id="8243f-172">名前</span><span class="sxs-lookup"><span data-stu-id="8243f-172">Name</span></span>| <span data-ttu-id="8243f-173">種類</span><span class="sxs-lookup"><span data-stu-id="8243f-173">Type</span></span>| <span data-ttu-id="8243f-174">説明</span><span class="sxs-lookup"><span data-stu-id="8243f-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="8243f-175">文字列</span><span class="sxs-lookup"><span data-stu-id="8243f-175">String</span></span>|<span data-ttu-id="8243f-176">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="8243f-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="8243f-177">String</span><span class="sxs-lookup"><span data-stu-id="8243f-177">String</span></span>|<span data-ttu-id="8243f-178">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="8243f-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8243f-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="8243f-179">Requirements</span></span>

|<span data-ttu-id="8243f-180">要件</span><span class="sxs-lookup"><span data-stu-id="8243f-180">Requirement</span></span>| <span data-ttu-id="8243f-181">値</span><span class="sxs-lookup"><span data-stu-id="8243f-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="8243f-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8243f-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8243f-183">1.1</span><span class="sxs-lookup"><span data-stu-id="8243f-183">1.1</span></span>|
|[<span data-ttu-id="8243f-184">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8243f-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8243f-185">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8243f-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="8243f-186">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="8243f-186">SourceProperty: String</span></span>

<span data-ttu-id="8243f-187">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="8243f-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8243f-188">型</span><span class="sxs-lookup"><span data-stu-id="8243f-188">Type</span></span>

*   <span data-ttu-id="8243f-189">String</span><span class="sxs-lookup"><span data-stu-id="8243f-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8243f-190">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="8243f-190">Properties:</span></span>

|<span data-ttu-id="8243f-191">名前</span><span class="sxs-lookup"><span data-stu-id="8243f-191">Name</span></span>| <span data-ttu-id="8243f-192">種類</span><span class="sxs-lookup"><span data-stu-id="8243f-192">Type</span></span>| <span data-ttu-id="8243f-193">説明</span><span class="sxs-lookup"><span data-stu-id="8243f-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="8243f-194">文字列</span><span class="sxs-lookup"><span data-stu-id="8243f-194">String</span></span>|<span data-ttu-id="8243f-195">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="8243f-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="8243f-196">文字列</span><span class="sxs-lookup"><span data-stu-id="8243f-196">String</span></span>|<span data-ttu-id="8243f-197">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="8243f-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8243f-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="8243f-198">Requirements</span></span>

|<span data-ttu-id="8243f-199">要件</span><span class="sxs-lookup"><span data-stu-id="8243f-199">Requirement</span></span>| <span data-ttu-id="8243f-200">値</span><span class="sxs-lookup"><span data-stu-id="8243f-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="8243f-201">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8243f-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8243f-202">1.1</span><span class="sxs-lookup"><span data-stu-id="8243f-202">1.1</span></span>|
|[<span data-ttu-id="8243f-203">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8243f-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8243f-204">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8243f-204">Compose or Read</span></span>|