---
title: Office 名前空間-要件セット1.6
description: メールボックス API 要件セット1.6 を使用した Outlook アドインで使用可能な Office 名前空間メンバー。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: dc7f62cc3f01e56f6c05b6cf40a4b73e87aea5e4
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891314"
---
# <a name="office-mailbox-requirement-set-16"></a><span data-ttu-id="3f372-103">Office (メールボックス要件セット 1.6)</span><span class="sxs-lookup"><span data-stu-id="3f372-103">Office (Mailbox requirement set 1.6)</span></span>

<span data-ttu-id="3f372-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3f372-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="3f372-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="3f372-106">Requirements</span></span>

|<span data-ttu-id="3f372-107">要件</span><span class="sxs-lookup"><span data-stu-id="3f372-107">Requirement</span></span>| <span data-ttu-id="3f372-108">値</span><span class="sxs-lookup"><span data-stu-id="3f372-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f372-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3f372-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3f372-110">1.1</span><span class="sxs-lookup"><span data-stu-id="3f372-110">1.1</span></span>|
|[<span data-ttu-id="3f372-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3f372-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3f372-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3f372-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="3f372-113">Properties</span><span class="sxs-lookup"><span data-stu-id="3f372-113">Properties</span></span>

| <span data-ttu-id="3f372-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="3f372-114">Property</span></span> | <span data-ttu-id="3f372-115">モード</span><span class="sxs-lookup"><span data-stu-id="3f372-115">Modes</span></span> | <span data-ttu-id="3f372-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="3f372-116">Return type</span></span> | <span data-ttu-id="3f372-117">最小値</span><span class="sxs-lookup"><span data-stu-id="3f372-117">Minimum</span></span><br><span data-ttu-id="3f372-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="3f372-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="3f372-119">context</span><span class="sxs-lookup"><span data-stu-id="3f372-119">context</span></span>](office.context.md) | <span data-ttu-id="3f372-120">作成</span><span class="sxs-lookup"><span data-stu-id="3f372-120">Compose</span></span><br><span data-ttu-id="3f372-121">読み取り</span><span class="sxs-lookup"><span data-stu-id="3f372-121">Read</span></span> | [<span data-ttu-id="3f372-122">Context</span><span class="sxs-lookup"><span data-stu-id="3f372-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6) | [<span data-ttu-id="3f372-123">1.1</span><span class="sxs-lookup"><span data-stu-id="3f372-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="3f372-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="3f372-124">Enumerations</span></span>

| <span data-ttu-id="3f372-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="3f372-125">Enumeration</span></span> | <span data-ttu-id="3f372-126">モード</span><span class="sxs-lookup"><span data-stu-id="3f372-126">Modes</span></span> | <span data-ttu-id="3f372-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="3f372-127">Return type</span></span> | <span data-ttu-id="3f372-128">最小値</span><span class="sxs-lookup"><span data-stu-id="3f372-128">Minimum</span></span><br><span data-ttu-id="3f372-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="3f372-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="3f372-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="3f372-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="3f372-131">作成</span><span class="sxs-lookup"><span data-stu-id="3f372-131">Compose</span></span><br><span data-ttu-id="3f372-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="3f372-132">Read</span></span> | <span data-ttu-id="3f372-133">String</span><span class="sxs-lookup"><span data-stu-id="3f372-133">String</span></span> | [<span data-ttu-id="3f372-134">1.1</span><span class="sxs-lookup"><span data-stu-id="3f372-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3f372-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="3f372-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="3f372-136">作成</span><span class="sxs-lookup"><span data-stu-id="3f372-136">Compose</span></span><br><span data-ttu-id="3f372-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="3f372-137">Read</span></span> | <span data-ttu-id="3f372-138">String</span><span class="sxs-lookup"><span data-stu-id="3f372-138">String</span></span> | [<span data-ttu-id="3f372-139">1.1</span><span class="sxs-lookup"><span data-stu-id="3f372-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3f372-140">EventType</span><span class="sxs-lookup"><span data-stu-id="3f372-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="3f372-141">作成</span><span class="sxs-lookup"><span data-stu-id="3f372-141">Compose</span></span><br><span data-ttu-id="3f372-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="3f372-142">Read</span></span> | <span data-ttu-id="3f372-143">String</span><span class="sxs-lookup"><span data-stu-id="3f372-143">String</span></span> | [<span data-ttu-id="3f372-144">1.5</span><span class="sxs-lookup"><span data-stu-id="3f372-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="3f372-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="3f372-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="3f372-146">作成</span><span class="sxs-lookup"><span data-stu-id="3f372-146">Compose</span></span><br><span data-ttu-id="3f372-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="3f372-147">Read</span></span> | <span data-ttu-id="3f372-148">String</span><span class="sxs-lookup"><span data-stu-id="3f372-148">String</span></span> | [<span data-ttu-id="3f372-149">1.1</span><span class="sxs-lookup"><span data-stu-id="3f372-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="3f372-150">名前空間</span><span class="sxs-lookup"><span data-stu-id="3f372-150">Namespaces</span></span>

<span data-ttu-id="3f372-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="3f372-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="3f372-152">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="3f372-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="3f372-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="3f372-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="3f372-154">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="3f372-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="3f372-155">型</span><span class="sxs-lookup"><span data-stu-id="3f372-155">Type</span></span>

*   <span data-ttu-id="3f372-156">String</span><span class="sxs-lookup"><span data-stu-id="3f372-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3f372-157">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="3f372-157">Properties:</span></span>

|<span data-ttu-id="3f372-158">名前</span><span class="sxs-lookup"><span data-stu-id="3f372-158">Name</span></span>| <span data-ttu-id="3f372-159">種類</span><span class="sxs-lookup"><span data-stu-id="3f372-159">Type</span></span>| <span data-ttu-id="3f372-160">説明</span><span class="sxs-lookup"><span data-stu-id="3f372-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="3f372-161">String</span><span class="sxs-lookup"><span data-stu-id="3f372-161">String</span></span>|<span data-ttu-id="3f372-162">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="3f372-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="3f372-163">String</span><span class="sxs-lookup"><span data-stu-id="3f372-163">String</span></span>|<span data-ttu-id="3f372-164">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="3f372-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3f372-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="3f372-165">Requirements</span></span>

|<span data-ttu-id="3f372-166">要件</span><span class="sxs-lookup"><span data-stu-id="3f372-166">Requirement</span></span>| <span data-ttu-id="3f372-167">値</span><span class="sxs-lookup"><span data-stu-id="3f372-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f372-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3f372-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3f372-169">1.1</span><span class="sxs-lookup"><span data-stu-id="3f372-169">1.1</span></span>|
|[<span data-ttu-id="3f372-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3f372-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3f372-171">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3f372-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="3f372-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="3f372-172">CoercionType: String</span></span>

<span data-ttu-id="3f372-173">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="3f372-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3f372-174">型</span><span class="sxs-lookup"><span data-stu-id="3f372-174">Type</span></span>

*   <span data-ttu-id="3f372-175">String</span><span class="sxs-lookup"><span data-stu-id="3f372-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3f372-176">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="3f372-176">Properties:</span></span>

|<span data-ttu-id="3f372-177">名前</span><span class="sxs-lookup"><span data-stu-id="3f372-177">Name</span></span>| <span data-ttu-id="3f372-178">種類</span><span class="sxs-lookup"><span data-stu-id="3f372-178">Type</span></span>| <span data-ttu-id="3f372-179">説明</span><span class="sxs-lookup"><span data-stu-id="3f372-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="3f372-180">String</span><span class="sxs-lookup"><span data-stu-id="3f372-180">String</span></span>|<span data-ttu-id="3f372-181">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="3f372-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="3f372-182">String</span><span class="sxs-lookup"><span data-stu-id="3f372-182">String</span></span>|<span data-ttu-id="3f372-183">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="3f372-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3f372-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="3f372-184">Requirements</span></span>

|<span data-ttu-id="3f372-185">要件</span><span class="sxs-lookup"><span data-stu-id="3f372-185">Requirement</span></span>| <span data-ttu-id="3f372-186">値</span><span class="sxs-lookup"><span data-stu-id="3f372-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f372-187">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3f372-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3f372-188">1.1</span><span class="sxs-lookup"><span data-stu-id="3f372-188">1.1</span></span>|
|[<span data-ttu-id="3f372-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3f372-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3f372-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3f372-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="3f372-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="3f372-191">EventType: String</span></span>

<span data-ttu-id="3f372-192">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="3f372-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="3f372-193">型</span><span class="sxs-lookup"><span data-stu-id="3f372-193">Type</span></span>

*   <span data-ttu-id="3f372-194">String</span><span class="sxs-lookup"><span data-stu-id="3f372-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3f372-195">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="3f372-195">Properties:</span></span>

| <span data-ttu-id="3f372-196">名前</span><span class="sxs-lookup"><span data-stu-id="3f372-196">Name</span></span> | <span data-ttu-id="3f372-197">種類</span><span class="sxs-lookup"><span data-stu-id="3f372-197">Type</span></span> | <span data-ttu-id="3f372-198">説明</span><span class="sxs-lookup"><span data-stu-id="3f372-198">Description</span></span> | <span data-ttu-id="3f372-199">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="3f372-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="3f372-200">String</span><span class="sxs-lookup"><span data-stu-id="3f372-200">String</span></span> | <span data-ttu-id="3f372-201">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="3f372-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="3f372-202">1.5</span><span class="sxs-lookup"><span data-stu-id="3f372-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3f372-203">Requirements</span><span class="sxs-lookup"><span data-stu-id="3f372-203">Requirements</span></span>

|<span data-ttu-id="3f372-204">要件</span><span class="sxs-lookup"><span data-stu-id="3f372-204">Requirement</span></span>| <span data-ttu-id="3f372-205">値</span><span class="sxs-lookup"><span data-stu-id="3f372-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f372-206">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3f372-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3f372-207">1.5</span><span class="sxs-lookup"><span data-stu-id="3f372-207">1.5</span></span> |
|[<span data-ttu-id="3f372-208">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3f372-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3f372-209">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3f372-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="3f372-210">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="3f372-210">SourceProperty: String</span></span>

<span data-ttu-id="3f372-211">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="3f372-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3f372-212">型</span><span class="sxs-lookup"><span data-stu-id="3f372-212">Type</span></span>

*   <span data-ttu-id="3f372-213">String</span><span class="sxs-lookup"><span data-stu-id="3f372-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3f372-214">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="3f372-214">Properties:</span></span>

|<span data-ttu-id="3f372-215">名前</span><span class="sxs-lookup"><span data-stu-id="3f372-215">Name</span></span>| <span data-ttu-id="3f372-216">種類</span><span class="sxs-lookup"><span data-stu-id="3f372-216">Type</span></span>| <span data-ttu-id="3f372-217">説明</span><span class="sxs-lookup"><span data-stu-id="3f372-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="3f372-218">String</span><span class="sxs-lookup"><span data-stu-id="3f372-218">String</span></span>|<span data-ttu-id="3f372-219">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="3f372-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="3f372-220">String</span><span class="sxs-lookup"><span data-stu-id="3f372-220">String</span></span>|<span data-ttu-id="3f372-221">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="3f372-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3f372-222">Requirements</span><span class="sxs-lookup"><span data-stu-id="3f372-222">Requirements</span></span>

|<span data-ttu-id="3f372-223">要件</span><span class="sxs-lookup"><span data-stu-id="3f372-223">Requirement</span></span>| <span data-ttu-id="3f372-224">値</span><span class="sxs-lookup"><span data-stu-id="3f372-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="3f372-225">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3f372-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3f372-226">1.1</span><span class="sxs-lookup"><span data-stu-id="3f372-226">1.1</span></span>|
|[<span data-ttu-id="3f372-227">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3f372-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3f372-228">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3f372-228">Compose or Read</span></span>|
