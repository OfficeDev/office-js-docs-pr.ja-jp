---
title: Office 名前空間-要件セット1.8
description: Office 名前空間は、Outlook Office アドインの共有インターフェイスを提供します (要件セット 1.8)
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 0bbe212b0b8e5dc1348cb5cdc03509c44a716d1a
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717503"
---
# <a name="office"></a><span data-ttu-id="863aa-103">Office</span><span class="sxs-lookup"><span data-stu-id="863aa-103">Office</span></span>

<span data-ttu-id="863aa-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="863aa-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="863aa-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="863aa-106">Requirements</span></span>

|<span data-ttu-id="863aa-107">要件</span><span class="sxs-lookup"><span data-stu-id="863aa-107">Requirement</span></span>| <span data-ttu-id="863aa-108">値</span><span class="sxs-lookup"><span data-stu-id="863aa-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="863aa-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="863aa-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="863aa-110">1.1</span><span class="sxs-lookup"><span data-stu-id="863aa-110">1.1</span></span>|
|[<span data-ttu-id="863aa-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="863aa-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="863aa-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="863aa-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="863aa-113">Properties</span><span class="sxs-lookup"><span data-stu-id="863aa-113">Properties</span></span>

| <span data-ttu-id="863aa-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="863aa-114">Property</span></span> | <span data-ttu-id="863aa-115">モード</span><span class="sxs-lookup"><span data-stu-id="863aa-115">Modes</span></span> | <span data-ttu-id="863aa-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="863aa-116">Return type</span></span> | <span data-ttu-id="863aa-117">最小値</span><span class="sxs-lookup"><span data-stu-id="863aa-117">Minimum</span></span><br><span data-ttu-id="863aa-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="863aa-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="863aa-119">context</span><span class="sxs-lookup"><span data-stu-id="863aa-119">context</span></span>](office.context.md) | <span data-ttu-id="863aa-120">作成</span><span class="sxs-lookup"><span data-stu-id="863aa-120">Compose</span></span><br><span data-ttu-id="863aa-121">読み取り</span><span class="sxs-lookup"><span data-stu-id="863aa-121">Read</span></span> | [<span data-ttu-id="863aa-122">Context</span><span class="sxs-lookup"><span data-stu-id="863aa-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8) | [<span data-ttu-id="863aa-123">1.1</span><span class="sxs-lookup"><span data-stu-id="863aa-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="863aa-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="863aa-124">Enumerations</span></span>

| <span data-ttu-id="863aa-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="863aa-125">Enumeration</span></span> | <span data-ttu-id="863aa-126">モード</span><span class="sxs-lookup"><span data-stu-id="863aa-126">Modes</span></span> | <span data-ttu-id="863aa-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="863aa-127">Return type</span></span> | <span data-ttu-id="863aa-128">最小値</span><span class="sxs-lookup"><span data-stu-id="863aa-128">Minimum</span></span><br><span data-ttu-id="863aa-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="863aa-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="863aa-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="863aa-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="863aa-131">作成</span><span class="sxs-lookup"><span data-stu-id="863aa-131">Compose</span></span><br><span data-ttu-id="863aa-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="863aa-132">Read</span></span> | <span data-ttu-id="863aa-133">文字列</span><span class="sxs-lookup"><span data-stu-id="863aa-133">String</span></span> | [<span data-ttu-id="863aa-134">1.1</span><span class="sxs-lookup"><span data-stu-id="863aa-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="863aa-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="863aa-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="863aa-136">作成</span><span class="sxs-lookup"><span data-stu-id="863aa-136">Compose</span></span><br><span data-ttu-id="863aa-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="863aa-137">Read</span></span> | <span data-ttu-id="863aa-138">文字列</span><span class="sxs-lookup"><span data-stu-id="863aa-138">String</span></span> | [<span data-ttu-id="863aa-139">1.1</span><span class="sxs-lookup"><span data-stu-id="863aa-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="863aa-140">EventType</span><span class="sxs-lookup"><span data-stu-id="863aa-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="863aa-141">作成</span><span class="sxs-lookup"><span data-stu-id="863aa-141">Compose</span></span><br><span data-ttu-id="863aa-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="863aa-142">Read</span></span> | <span data-ttu-id="863aa-143">文字列</span><span class="sxs-lookup"><span data-stu-id="863aa-143">String</span></span> | [<span data-ttu-id="863aa-144">1.5</span><span class="sxs-lookup"><span data-stu-id="863aa-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="863aa-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="863aa-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="863aa-146">作成</span><span class="sxs-lookup"><span data-stu-id="863aa-146">Compose</span></span><br><span data-ttu-id="863aa-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="863aa-147">Read</span></span> | <span data-ttu-id="863aa-148">文字列</span><span class="sxs-lookup"><span data-stu-id="863aa-148">String</span></span> | [<span data-ttu-id="863aa-149">1.1</span><span class="sxs-lookup"><span data-stu-id="863aa-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="863aa-150">名前空間</span><span class="sxs-lookup"><span data-stu-id="863aa-150">Namespaces</span></span>

<span data-ttu-id="863aa-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="863aa-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="863aa-152">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="863aa-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="863aa-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="863aa-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="863aa-154">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="863aa-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="863aa-155">型</span><span class="sxs-lookup"><span data-stu-id="863aa-155">Type</span></span>

*   <span data-ttu-id="863aa-156">String</span><span class="sxs-lookup"><span data-stu-id="863aa-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="863aa-157">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="863aa-157">Properties:</span></span>

|<span data-ttu-id="863aa-158">名前</span><span class="sxs-lookup"><span data-stu-id="863aa-158">Name</span></span>| <span data-ttu-id="863aa-159">種類</span><span class="sxs-lookup"><span data-stu-id="863aa-159">Type</span></span>| <span data-ttu-id="863aa-160">説明</span><span class="sxs-lookup"><span data-stu-id="863aa-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="863aa-161">文字列</span><span class="sxs-lookup"><span data-stu-id="863aa-161">String</span></span>|<span data-ttu-id="863aa-162">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="863aa-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="863aa-163">文字列</span><span class="sxs-lookup"><span data-stu-id="863aa-163">String</span></span>|<span data-ttu-id="863aa-164">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="863aa-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="863aa-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="863aa-165">Requirements</span></span>

|<span data-ttu-id="863aa-166">要件</span><span class="sxs-lookup"><span data-stu-id="863aa-166">Requirement</span></span>| <span data-ttu-id="863aa-167">値</span><span class="sxs-lookup"><span data-stu-id="863aa-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="863aa-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="863aa-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="863aa-169">1.1</span><span class="sxs-lookup"><span data-stu-id="863aa-169">1.1</span></span>|
|[<span data-ttu-id="863aa-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="863aa-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="863aa-171">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="863aa-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="863aa-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="863aa-172">CoercionType: String</span></span>

<span data-ttu-id="863aa-173">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="863aa-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="863aa-174">型</span><span class="sxs-lookup"><span data-stu-id="863aa-174">Type</span></span>

*   <span data-ttu-id="863aa-175">String</span><span class="sxs-lookup"><span data-stu-id="863aa-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="863aa-176">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="863aa-176">Properties:</span></span>

|<span data-ttu-id="863aa-177">名前</span><span class="sxs-lookup"><span data-stu-id="863aa-177">Name</span></span>| <span data-ttu-id="863aa-178">種類</span><span class="sxs-lookup"><span data-stu-id="863aa-178">Type</span></span>| <span data-ttu-id="863aa-179">説明</span><span class="sxs-lookup"><span data-stu-id="863aa-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="863aa-180">文字列</span><span class="sxs-lookup"><span data-stu-id="863aa-180">String</span></span>|<span data-ttu-id="863aa-181">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="863aa-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="863aa-182">String</span><span class="sxs-lookup"><span data-stu-id="863aa-182">String</span></span>|<span data-ttu-id="863aa-183">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="863aa-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="863aa-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="863aa-184">Requirements</span></span>

|<span data-ttu-id="863aa-185">要件</span><span class="sxs-lookup"><span data-stu-id="863aa-185">Requirement</span></span>| <span data-ttu-id="863aa-186">値</span><span class="sxs-lookup"><span data-stu-id="863aa-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="863aa-187">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="863aa-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="863aa-188">1.1</span><span class="sxs-lookup"><span data-stu-id="863aa-188">1.1</span></span>|
|[<span data-ttu-id="863aa-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="863aa-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="863aa-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="863aa-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="863aa-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="863aa-191">EventType: String</span></span>

<span data-ttu-id="863aa-192">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="863aa-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="863aa-193">型</span><span class="sxs-lookup"><span data-stu-id="863aa-193">Type</span></span>

*   <span data-ttu-id="863aa-194">String</span><span class="sxs-lookup"><span data-stu-id="863aa-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="863aa-195">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="863aa-195">Properties:</span></span>

| <span data-ttu-id="863aa-196">名前</span><span class="sxs-lookup"><span data-stu-id="863aa-196">Name</span></span> | <span data-ttu-id="863aa-197">種類</span><span class="sxs-lookup"><span data-stu-id="863aa-197">Type</span></span> | <span data-ttu-id="863aa-198">説明</span><span class="sxs-lookup"><span data-stu-id="863aa-198">Description</span></span> | <span data-ttu-id="863aa-199">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="863aa-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="863aa-200">文字列</span><span class="sxs-lookup"><span data-stu-id="863aa-200">String</span></span> | <span data-ttu-id="863aa-201">選択した予定またはデータ系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="863aa-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="863aa-202">1.7</span><span class="sxs-lookup"><span data-stu-id="863aa-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="863aa-203">文字列</span><span class="sxs-lookup"><span data-stu-id="863aa-203">String</span></span> | <span data-ttu-id="863aa-204">添付ファイルがアイテムに追加またはアイテムから削除されています。</span><span class="sxs-lookup"><span data-stu-id="863aa-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="863aa-205">1.8</span><span class="sxs-lookup"><span data-stu-id="863aa-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="863aa-206">文字列</span><span class="sxs-lookup"><span data-stu-id="863aa-206">String</span></span> | <span data-ttu-id="863aa-207">選択した予定の場所が変更されました。</span><span class="sxs-lookup"><span data-stu-id="863aa-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="863aa-208">1.8</span><span class="sxs-lookup"><span data-stu-id="863aa-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="863aa-209">文字列</span><span class="sxs-lookup"><span data-stu-id="863aa-209">String</span></span> | <span data-ttu-id="863aa-210">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="863aa-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="863aa-211">1.5</span><span class="sxs-lookup"><span data-stu-id="863aa-211">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="863aa-212">文字列</span><span class="sxs-lookup"><span data-stu-id="863aa-212">String</span></span> | <span data-ttu-id="863aa-213">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="863aa-213">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="863aa-214">1.7</span><span class="sxs-lookup"><span data-stu-id="863aa-214">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="863aa-215">文字列</span><span class="sxs-lookup"><span data-stu-id="863aa-215">String</span></span> | <span data-ttu-id="863aa-216">選択したアイテムの定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="863aa-216">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="863aa-217">1.7</span><span class="sxs-lookup"><span data-stu-id="863aa-217">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="863aa-218">Requirements</span><span class="sxs-lookup"><span data-stu-id="863aa-218">Requirements</span></span>

|<span data-ttu-id="863aa-219">要件</span><span class="sxs-lookup"><span data-stu-id="863aa-219">Requirement</span></span>| <span data-ttu-id="863aa-220">値</span><span class="sxs-lookup"><span data-stu-id="863aa-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="863aa-221">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="863aa-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="863aa-222">1.5</span><span class="sxs-lookup"><span data-stu-id="863aa-222">1.5</span></span> |
|[<span data-ttu-id="863aa-223">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="863aa-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="863aa-224">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="863aa-224">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="863aa-225">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="863aa-225">SourceProperty: String</span></span>

<span data-ttu-id="863aa-226">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="863aa-226">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="863aa-227">型</span><span class="sxs-lookup"><span data-stu-id="863aa-227">Type</span></span>

*   <span data-ttu-id="863aa-228">String</span><span class="sxs-lookup"><span data-stu-id="863aa-228">String</span></span>

##### <a name="properties"></a><span data-ttu-id="863aa-229">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="863aa-229">Properties:</span></span>

|<span data-ttu-id="863aa-230">名前</span><span class="sxs-lookup"><span data-stu-id="863aa-230">Name</span></span>| <span data-ttu-id="863aa-231">種類</span><span class="sxs-lookup"><span data-stu-id="863aa-231">Type</span></span>| <span data-ttu-id="863aa-232">説明</span><span class="sxs-lookup"><span data-stu-id="863aa-232">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="863aa-233">文字列</span><span class="sxs-lookup"><span data-stu-id="863aa-233">String</span></span>|<span data-ttu-id="863aa-234">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="863aa-234">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="863aa-235">文字列</span><span class="sxs-lookup"><span data-stu-id="863aa-235">String</span></span>|<span data-ttu-id="863aa-236">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="863aa-236">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="863aa-237">Requirements</span><span class="sxs-lookup"><span data-stu-id="863aa-237">Requirements</span></span>

|<span data-ttu-id="863aa-238">要件</span><span class="sxs-lookup"><span data-stu-id="863aa-238">Requirement</span></span>| <span data-ttu-id="863aa-239">値</span><span class="sxs-lookup"><span data-stu-id="863aa-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="863aa-240">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="863aa-240">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="863aa-241">1.1</span><span class="sxs-lookup"><span data-stu-id="863aa-241">1.1</span></span>|
|[<span data-ttu-id="863aa-242">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="863aa-242">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="863aa-243">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="863aa-243">Compose or Read</span></span>|
