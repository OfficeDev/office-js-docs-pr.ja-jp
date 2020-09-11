---
title: Office 名前空間-要件セット1.4
description: メールボックス API 要件セット1.4 を使用した Outlook アドインで使用可能な Office 名前空間メンバー。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: f797fe5281d2031a2182249aeb18d740cd114d43
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430745"
---
# <a name="office-mailbox-requirement-set-14"></a><span data-ttu-id="641d4-103">Office (メールボックス要件セット 1.4)</span><span class="sxs-lookup"><span data-stu-id="641d4-103">Office (Mailbox requirement set 1.4)</span></span>

<span data-ttu-id="641d4-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="641d4-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="641d4-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="641d4-106">Requirements</span></span>

|<span data-ttu-id="641d4-107">要件</span><span class="sxs-lookup"><span data-stu-id="641d4-107">Requirement</span></span>| <span data-ttu-id="641d4-108">値</span><span class="sxs-lookup"><span data-stu-id="641d4-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="641d4-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="641d4-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="641d4-110">1.1</span><span class="sxs-lookup"><span data-stu-id="641d4-110">1.1</span></span>|
|[<span data-ttu-id="641d4-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="641d4-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="641d4-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="641d4-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="641d4-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="641d4-113">Properties</span></span>

| <span data-ttu-id="641d4-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="641d4-114">Property</span></span> | <span data-ttu-id="641d4-115">モード</span><span class="sxs-lookup"><span data-stu-id="641d4-115">Modes</span></span> | <span data-ttu-id="641d4-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="641d4-116">Return type</span></span> | <span data-ttu-id="641d4-117">最小値</span><span class="sxs-lookup"><span data-stu-id="641d4-117">Minimum</span></span><br><span data-ttu-id="641d4-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="641d4-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="641d4-119">context</span><span class="sxs-lookup"><span data-stu-id="641d4-119">context</span></span>](office.context.md) | <span data-ttu-id="641d4-120">作成</span><span class="sxs-lookup"><span data-stu-id="641d4-120">Compose</span></span><br><span data-ttu-id="641d4-121">読み取り</span><span class="sxs-lookup"><span data-stu-id="641d4-121">Read</span></span> | [<span data-ttu-id="641d4-122">Context</span><span class="sxs-lookup"><span data-stu-id="641d4-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="641d4-123">1.1</span><span class="sxs-lookup"><span data-stu-id="641d4-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="641d4-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="641d4-124">Enumerations</span></span>

| <span data-ttu-id="641d4-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="641d4-125">Enumeration</span></span> | <span data-ttu-id="641d4-126">モード</span><span class="sxs-lookup"><span data-stu-id="641d4-126">Modes</span></span> | <span data-ttu-id="641d4-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="641d4-127">Return type</span></span> | <span data-ttu-id="641d4-128">最小値</span><span class="sxs-lookup"><span data-stu-id="641d4-128">Minimum</span></span><br><span data-ttu-id="641d4-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="641d4-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="641d4-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="641d4-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="641d4-131">作成</span><span class="sxs-lookup"><span data-stu-id="641d4-131">Compose</span></span><br><span data-ttu-id="641d4-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="641d4-132">Read</span></span> | <span data-ttu-id="641d4-133">文字列</span><span class="sxs-lookup"><span data-stu-id="641d4-133">String</span></span> | [<span data-ttu-id="641d4-134">1.1</span><span class="sxs-lookup"><span data-stu-id="641d4-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="641d4-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="641d4-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="641d4-136">作成</span><span class="sxs-lookup"><span data-stu-id="641d4-136">Compose</span></span><br><span data-ttu-id="641d4-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="641d4-137">Read</span></span> | <span data-ttu-id="641d4-138">文字列</span><span class="sxs-lookup"><span data-stu-id="641d4-138">String</span></span> | [<span data-ttu-id="641d4-139">1.1</span><span class="sxs-lookup"><span data-stu-id="641d4-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="641d4-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="641d4-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="641d4-141">作成</span><span class="sxs-lookup"><span data-stu-id="641d4-141">Compose</span></span><br><span data-ttu-id="641d4-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="641d4-142">Read</span></span> | <span data-ttu-id="641d4-143">文字列</span><span class="sxs-lookup"><span data-stu-id="641d4-143">String</span></span> | [<span data-ttu-id="641d4-144">1.1</span><span class="sxs-lookup"><span data-stu-id="641d4-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="641d4-145">名前空間</span><span class="sxs-lookup"><span data-stu-id="641d4-145">Namespaces</span></span>

<span data-ttu-id="641d4-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4&preserve-view=true):、、、、、など、多数の Outlook 固有の列挙を含み `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` ます。</span><span class="sxs-lookup"><span data-stu-id="641d4-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="641d4-147">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="641d4-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="641d4-148">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="641d4-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="641d4-149">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="641d4-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="641d4-150">型</span><span class="sxs-lookup"><span data-stu-id="641d4-150">Type</span></span>

*   <span data-ttu-id="641d4-151">String</span><span class="sxs-lookup"><span data-stu-id="641d4-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="641d4-152">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="641d4-152">Properties:</span></span>

|<span data-ttu-id="641d4-153">名前</span><span class="sxs-lookup"><span data-stu-id="641d4-153">Name</span></span>| <span data-ttu-id="641d4-154">種類</span><span class="sxs-lookup"><span data-stu-id="641d4-154">Type</span></span>| <span data-ttu-id="641d4-155">説明</span><span class="sxs-lookup"><span data-stu-id="641d4-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="641d4-156">文字列</span><span class="sxs-lookup"><span data-stu-id="641d4-156">String</span></span>|<span data-ttu-id="641d4-157">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="641d4-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="641d4-158">String</span><span class="sxs-lookup"><span data-stu-id="641d4-158">String</span></span>|<span data-ttu-id="641d4-159">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="641d4-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="641d4-160">Requirements</span><span class="sxs-lookup"><span data-stu-id="641d4-160">Requirements</span></span>

|<span data-ttu-id="641d4-161">要件</span><span class="sxs-lookup"><span data-stu-id="641d4-161">Requirement</span></span>| <span data-ttu-id="641d4-162">値</span><span class="sxs-lookup"><span data-stu-id="641d4-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="641d4-163">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="641d4-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="641d4-164">1.1</span><span class="sxs-lookup"><span data-stu-id="641d4-164">1.1</span></span>|
|[<span data-ttu-id="641d4-165">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="641d4-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="641d4-166">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="641d4-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="641d4-167">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="641d4-167">CoercionType: String</span></span>

<span data-ttu-id="641d4-168">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="641d4-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="641d4-169">型</span><span class="sxs-lookup"><span data-stu-id="641d4-169">Type</span></span>

*   <span data-ttu-id="641d4-170">String</span><span class="sxs-lookup"><span data-stu-id="641d4-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="641d4-171">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="641d4-171">Properties:</span></span>

|<span data-ttu-id="641d4-172">名前</span><span class="sxs-lookup"><span data-stu-id="641d4-172">Name</span></span>| <span data-ttu-id="641d4-173">種類</span><span class="sxs-lookup"><span data-stu-id="641d4-173">Type</span></span>| <span data-ttu-id="641d4-174">説明</span><span class="sxs-lookup"><span data-stu-id="641d4-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="641d4-175">文字列</span><span class="sxs-lookup"><span data-stu-id="641d4-175">String</span></span>|<span data-ttu-id="641d4-176">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="641d4-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="641d4-177">String</span><span class="sxs-lookup"><span data-stu-id="641d4-177">String</span></span>|<span data-ttu-id="641d4-178">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="641d4-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="641d4-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="641d4-179">Requirements</span></span>

|<span data-ttu-id="641d4-180">要件</span><span class="sxs-lookup"><span data-stu-id="641d4-180">Requirement</span></span>| <span data-ttu-id="641d4-181">値</span><span class="sxs-lookup"><span data-stu-id="641d4-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="641d4-182">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="641d4-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="641d4-183">1.1</span><span class="sxs-lookup"><span data-stu-id="641d4-183">1.1</span></span>|
|[<span data-ttu-id="641d4-184">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="641d4-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="641d4-185">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="641d4-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="641d4-186">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="641d4-186">SourceProperty: String</span></span>

<span data-ttu-id="641d4-187">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="641d4-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="641d4-188">型</span><span class="sxs-lookup"><span data-stu-id="641d4-188">Type</span></span>

*   <span data-ttu-id="641d4-189">String</span><span class="sxs-lookup"><span data-stu-id="641d4-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="641d4-190">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="641d4-190">Properties:</span></span>

|<span data-ttu-id="641d4-191">名前</span><span class="sxs-lookup"><span data-stu-id="641d4-191">Name</span></span>| <span data-ttu-id="641d4-192">種類</span><span class="sxs-lookup"><span data-stu-id="641d4-192">Type</span></span>| <span data-ttu-id="641d4-193">説明</span><span class="sxs-lookup"><span data-stu-id="641d4-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="641d4-194">文字列</span><span class="sxs-lookup"><span data-stu-id="641d4-194">String</span></span>|<span data-ttu-id="641d4-195">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="641d4-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="641d4-196">String</span><span class="sxs-lookup"><span data-stu-id="641d4-196">String</span></span>|<span data-ttu-id="641d4-197">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="641d4-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="641d4-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="641d4-198">Requirements</span></span>

|<span data-ttu-id="641d4-199">要件</span><span class="sxs-lookup"><span data-stu-id="641d4-199">Requirement</span></span>| <span data-ttu-id="641d4-200">値</span><span class="sxs-lookup"><span data-stu-id="641d4-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="641d4-201">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="641d4-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="641d4-202">1.1</span><span class="sxs-lookup"><span data-stu-id="641d4-202">1.1</span></span>|
|[<span data-ttu-id="641d4-203">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="641d4-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="641d4-204">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="641d4-204">Compose or Read</span></span>|
