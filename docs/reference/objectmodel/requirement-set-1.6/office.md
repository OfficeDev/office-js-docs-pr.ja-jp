---
title: Office名前空間 - 要件セット 1.6
description: Office API 要件セット 1.6 をOutlookアドインで使用できる名前空間メンバーを指定します。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 40cdb7de0678007b93b9251e7f1e2921ed857338
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590835"
---
# <a name="office-mailbox-requirement-set-16"></a><span data-ttu-id="f11a5-103">Office (メールボックス要件セット 1.6)</span><span class="sxs-lookup"><span data-stu-id="f11a5-103">Office (Mailbox requirement set 1.6)</span></span>

<span data-ttu-id="f11a5-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f11a5-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f11a5-106">要件</span><span class="sxs-lookup"><span data-stu-id="f11a5-106">Requirements</span></span>

|<span data-ttu-id="f11a5-107">要件</span><span class="sxs-lookup"><span data-stu-id="f11a5-107">Requirement</span></span>| <span data-ttu-id="f11a5-108">値</span><span class="sxs-lookup"><span data-stu-id="f11a5-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="f11a5-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f11a5-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f11a5-110">1.1</span><span class="sxs-lookup"><span data-stu-id="f11a5-110">1.1</span></span>|
|[<span data-ttu-id="f11a5-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f11a5-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f11a5-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f11a5-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="f11a5-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="f11a5-113">Properties</span></span>

| <span data-ttu-id="f11a5-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="f11a5-114">Property</span></span> | <span data-ttu-id="f11a5-115">モード</span><span class="sxs-lookup"><span data-stu-id="f11a5-115">Modes</span></span> | <span data-ttu-id="f11a5-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="f11a5-116">Return type</span></span> | <span data-ttu-id="f11a5-117">最小値</span><span class="sxs-lookup"><span data-stu-id="f11a5-117">Minimum</span></span><br><span data-ttu-id="f11a5-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="f11a5-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f11a5-119">context</span><span class="sxs-lookup"><span data-stu-id="f11a5-119">context</span></span>](office.context.md) | <span data-ttu-id="f11a5-120">作成</span><span class="sxs-lookup"><span data-stu-id="f11a5-120">Compose</span></span><br><span data-ttu-id="f11a5-121">Read</span><span class="sxs-lookup"><span data-stu-id="f11a5-121">Read</span></span> | [<span data-ttu-id="f11a5-122">Context</span><span class="sxs-lookup"><span data-stu-id="f11a5-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="f11a5-123">1.1</span><span class="sxs-lookup"><span data-stu-id="f11a5-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="f11a5-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="f11a5-124">Enumerations</span></span>

| <span data-ttu-id="f11a5-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="f11a5-125">Enumeration</span></span> | <span data-ttu-id="f11a5-126">モード</span><span class="sxs-lookup"><span data-stu-id="f11a5-126">Modes</span></span> | <span data-ttu-id="f11a5-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="f11a5-127">Return type</span></span> | <span data-ttu-id="f11a5-128">最小値</span><span class="sxs-lookup"><span data-stu-id="f11a5-128">Minimum</span></span><br><span data-ttu-id="f11a5-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="f11a5-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f11a5-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="f11a5-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="f11a5-131">作成</span><span class="sxs-lookup"><span data-stu-id="f11a5-131">Compose</span></span><br><span data-ttu-id="f11a5-132">Read</span><span class="sxs-lookup"><span data-stu-id="f11a5-132">Read</span></span> | <span data-ttu-id="f11a5-133">String</span><span class="sxs-lookup"><span data-stu-id="f11a5-133">String</span></span> | [<span data-ttu-id="f11a5-134">1.1</span><span class="sxs-lookup"><span data-stu-id="f11a5-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f11a5-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="f11a5-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="f11a5-136">作成</span><span class="sxs-lookup"><span data-stu-id="f11a5-136">Compose</span></span><br><span data-ttu-id="f11a5-137">Read</span><span class="sxs-lookup"><span data-stu-id="f11a5-137">Read</span></span> | <span data-ttu-id="f11a5-138">String</span><span class="sxs-lookup"><span data-stu-id="f11a5-138">String</span></span> | [<span data-ttu-id="f11a5-139">1.1</span><span class="sxs-lookup"><span data-stu-id="f11a5-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f11a5-140">EventType</span><span class="sxs-lookup"><span data-stu-id="f11a5-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="f11a5-141">作成</span><span class="sxs-lookup"><span data-stu-id="f11a5-141">Compose</span></span><br><span data-ttu-id="f11a5-142">Read</span><span class="sxs-lookup"><span data-stu-id="f11a5-142">Read</span></span> | <span data-ttu-id="f11a5-143">String</span><span class="sxs-lookup"><span data-stu-id="f11a5-143">String</span></span> | [<span data-ttu-id="f11a5-144">1.5</span><span class="sxs-lookup"><span data-stu-id="f11a5-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="f11a5-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="f11a5-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="f11a5-146">作成</span><span class="sxs-lookup"><span data-stu-id="f11a5-146">Compose</span></span><br><span data-ttu-id="f11a5-147">Read</span><span class="sxs-lookup"><span data-stu-id="f11a5-147">Read</span></span> | <span data-ttu-id="f11a5-148">String</span><span class="sxs-lookup"><span data-stu-id="f11a5-148">String</span></span> | [<span data-ttu-id="f11a5-149">1.1</span><span class="sxs-lookup"><span data-stu-id="f11a5-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="f11a5-150">名前空間</span><span class="sxs-lookup"><span data-stu-id="f11a5-150">Namespaces</span></span>

<span data-ttu-id="f11a5-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6&preserve-view=true): 、 など、Outlook固有の列挙の `ItemType` `EntityType` `AttachmentType` 数 `RecipientType` が `ResponseType` 含まれています `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="f11a5-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="f11a5-152">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="f11a5-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="f11a5-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="f11a5-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="f11a5-154">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="f11a5-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f11a5-155">型</span><span class="sxs-lookup"><span data-stu-id="f11a5-155">Type</span></span>

*   <span data-ttu-id="f11a5-156">String</span><span class="sxs-lookup"><span data-stu-id="f11a5-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f11a5-157">プロパティ</span><span class="sxs-lookup"><span data-stu-id="f11a5-157">Properties</span></span>

|<span data-ttu-id="f11a5-158">名前</span><span class="sxs-lookup"><span data-stu-id="f11a5-158">Name</span></span>| <span data-ttu-id="f11a5-159">型</span><span class="sxs-lookup"><span data-stu-id="f11a5-159">Type</span></span>| <span data-ttu-id="f11a5-160">説明</span><span class="sxs-lookup"><span data-stu-id="f11a5-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f11a5-161">String</span><span class="sxs-lookup"><span data-stu-id="f11a5-161">String</span></span>|<span data-ttu-id="f11a5-162">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="f11a5-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f11a5-163">String</span><span class="sxs-lookup"><span data-stu-id="f11a5-163">String</span></span>|<span data-ttu-id="f11a5-164">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="f11a5-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f11a5-165">要件</span><span class="sxs-lookup"><span data-stu-id="f11a5-165">Requirements</span></span>

|<span data-ttu-id="f11a5-166">要件</span><span class="sxs-lookup"><span data-stu-id="f11a5-166">Requirement</span></span>| <span data-ttu-id="f11a5-167">値</span><span class="sxs-lookup"><span data-stu-id="f11a5-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="f11a5-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f11a5-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f11a5-169">1.1</span><span class="sxs-lookup"><span data-stu-id="f11a5-169">1.1</span></span>|
|[<span data-ttu-id="f11a5-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f11a5-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f11a5-171">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f11a5-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="f11a5-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="f11a5-172">CoercionType: String</span></span>

<span data-ttu-id="f11a5-173">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="f11a5-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f11a5-174">型</span><span class="sxs-lookup"><span data-stu-id="f11a5-174">Type</span></span>

*   <span data-ttu-id="f11a5-175">String</span><span class="sxs-lookup"><span data-stu-id="f11a5-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f11a5-176">プロパティ</span><span class="sxs-lookup"><span data-stu-id="f11a5-176">Properties</span></span>

|<span data-ttu-id="f11a5-177">名前</span><span class="sxs-lookup"><span data-stu-id="f11a5-177">Name</span></span>| <span data-ttu-id="f11a5-178">型</span><span class="sxs-lookup"><span data-stu-id="f11a5-178">Type</span></span>| <span data-ttu-id="f11a5-179">説明</span><span class="sxs-lookup"><span data-stu-id="f11a5-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f11a5-180">String</span><span class="sxs-lookup"><span data-stu-id="f11a5-180">String</span></span>|<span data-ttu-id="f11a5-181">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="f11a5-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f11a5-182">String</span><span class="sxs-lookup"><span data-stu-id="f11a5-182">String</span></span>|<span data-ttu-id="f11a5-183">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="f11a5-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f11a5-184">要件</span><span class="sxs-lookup"><span data-stu-id="f11a5-184">Requirements</span></span>

|<span data-ttu-id="f11a5-185">要件</span><span class="sxs-lookup"><span data-stu-id="f11a5-185">Requirement</span></span>| <span data-ttu-id="f11a5-186">値</span><span class="sxs-lookup"><span data-stu-id="f11a5-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="f11a5-187">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f11a5-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f11a5-188">1.1</span><span class="sxs-lookup"><span data-stu-id="f11a5-188">1.1</span></span>|
|[<span data-ttu-id="f11a5-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f11a5-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f11a5-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f11a5-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="f11a5-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="f11a5-191">EventType: String</span></span>

<span data-ttu-id="f11a5-192">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="f11a5-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="f11a5-193">型</span><span class="sxs-lookup"><span data-stu-id="f11a5-193">Type</span></span>

*   <span data-ttu-id="f11a5-194">String</span><span class="sxs-lookup"><span data-stu-id="f11a5-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f11a5-195">プロパティ</span><span class="sxs-lookup"><span data-stu-id="f11a5-195">Properties</span></span>

| <span data-ttu-id="f11a5-196">名前</span><span class="sxs-lookup"><span data-stu-id="f11a5-196">Name</span></span> | <span data-ttu-id="f11a5-197">型</span><span class="sxs-lookup"><span data-stu-id="f11a5-197">Type</span></span> | <span data-ttu-id="f11a5-198">説明</span><span class="sxs-lookup"><span data-stu-id="f11a5-198">Description</span></span> | <span data-ttu-id="f11a5-199">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="f11a5-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="f11a5-200">String</span><span class="sxs-lookup"><span data-stu-id="f11a5-200">String</span></span> | <span data-ttu-id="f11a5-201">作業ウィンドウOutlook表示する場合は、別のアイテムが選択されています。</span><span class="sxs-lookup"><span data-stu-id="f11a5-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="f11a5-202">1.5</span><span class="sxs-lookup"><span data-stu-id="f11a5-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f11a5-203">要件</span><span class="sxs-lookup"><span data-stu-id="f11a5-203">Requirements</span></span>

|<span data-ttu-id="f11a5-204">要件</span><span class="sxs-lookup"><span data-stu-id="f11a5-204">Requirement</span></span>| <span data-ttu-id="f11a5-205">値</span><span class="sxs-lookup"><span data-stu-id="f11a5-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="f11a5-206">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f11a5-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f11a5-207">1.5</span><span class="sxs-lookup"><span data-stu-id="f11a5-207">1.5</span></span> |
|[<span data-ttu-id="f11a5-208">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f11a5-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f11a5-209">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f11a5-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="f11a5-210">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="f11a5-210">SourceProperty: String</span></span>

<span data-ttu-id="f11a5-211">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="f11a5-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f11a5-212">型</span><span class="sxs-lookup"><span data-stu-id="f11a5-212">Type</span></span>

*   <span data-ttu-id="f11a5-213">String</span><span class="sxs-lookup"><span data-stu-id="f11a5-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f11a5-214">プロパティ</span><span class="sxs-lookup"><span data-stu-id="f11a5-214">Properties</span></span>

|<span data-ttu-id="f11a5-215">名前</span><span class="sxs-lookup"><span data-stu-id="f11a5-215">Name</span></span>| <span data-ttu-id="f11a5-216">型</span><span class="sxs-lookup"><span data-stu-id="f11a5-216">Type</span></span>| <span data-ttu-id="f11a5-217">説明</span><span class="sxs-lookup"><span data-stu-id="f11a5-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f11a5-218">String</span><span class="sxs-lookup"><span data-stu-id="f11a5-218">String</span></span>|<span data-ttu-id="f11a5-219">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="f11a5-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f11a5-220">String</span><span class="sxs-lookup"><span data-stu-id="f11a5-220">String</span></span>|<span data-ttu-id="f11a5-221">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="f11a5-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f11a5-222">要件</span><span class="sxs-lookup"><span data-stu-id="f11a5-222">Requirements</span></span>

|<span data-ttu-id="f11a5-223">要件</span><span class="sxs-lookup"><span data-stu-id="f11a5-223">Requirement</span></span>| <span data-ttu-id="f11a5-224">値</span><span class="sxs-lookup"><span data-stu-id="f11a5-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="f11a5-225">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f11a5-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f11a5-226">1.1</span><span class="sxs-lookup"><span data-stu-id="f11a5-226">1.1</span></span>|
|[<span data-ttu-id="f11a5-227">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f11a5-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f11a5-228">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f11a5-228">Compose or Read</span></span>|
