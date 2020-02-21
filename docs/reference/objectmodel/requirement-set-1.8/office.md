---
title: Office 名前空間-要件セット1.8
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: c5c431f7a958f1c2a956f36e90ad0f3a205c6669
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163627"
---
# <a name="office"></a><span data-ttu-id="e884f-102">Office</span><span class="sxs-lookup"><span data-stu-id="e884f-102">Office</span></span>

<span data-ttu-id="e884f-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e884f-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e884f-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="e884f-105">Requirements</span></span>

|<span data-ttu-id="e884f-106">要件</span><span class="sxs-lookup"><span data-stu-id="e884f-106">Requirement</span></span>| <span data-ttu-id="e884f-107">値</span><span class="sxs-lookup"><span data-stu-id="e884f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e884f-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e884f-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e884f-109">1.1</span><span class="sxs-lookup"><span data-stu-id="e884f-109">1.1</span></span>|
|[<span data-ttu-id="e884f-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e884f-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e884f-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e884f-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="e884f-112">Properties</span><span class="sxs-lookup"><span data-stu-id="e884f-112">Properties</span></span>

| <span data-ttu-id="e884f-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="e884f-113">Property</span></span> | <span data-ttu-id="e884f-114">モード</span><span class="sxs-lookup"><span data-stu-id="e884f-114">Modes</span></span> | <span data-ttu-id="e884f-115">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="e884f-115">Return type</span></span> | <span data-ttu-id="e884f-116">最小値</span><span class="sxs-lookup"><span data-stu-id="e884f-116">Minimum</span></span><br><span data-ttu-id="e884f-117">要件セット</span><span class="sxs-lookup"><span data-stu-id="e884f-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e884f-118">context</span><span class="sxs-lookup"><span data-stu-id="e884f-118">context</span></span>](office.context.md) | <span data-ttu-id="e884f-119">作成</span><span class="sxs-lookup"><span data-stu-id="e884f-119">Compose</span></span><br><span data-ttu-id="e884f-120">読み取り</span><span class="sxs-lookup"><span data-stu-id="e884f-120">Read</span></span> | [<span data-ttu-id="e884f-121">Context</span><span class="sxs-lookup"><span data-stu-id="e884f-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8) | [<span data-ttu-id="e884f-122">1.1</span><span class="sxs-lookup"><span data-stu-id="e884f-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="e884f-123">列挙型</span><span class="sxs-lookup"><span data-stu-id="e884f-123">Enumerations</span></span>

| <span data-ttu-id="e884f-124">列挙体</span><span class="sxs-lookup"><span data-stu-id="e884f-124">Enumeration</span></span> | <span data-ttu-id="e884f-125">モード</span><span class="sxs-lookup"><span data-stu-id="e884f-125">Modes</span></span> | <span data-ttu-id="e884f-126">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="e884f-126">Return type</span></span> | <span data-ttu-id="e884f-127">最小値</span><span class="sxs-lookup"><span data-stu-id="e884f-127">Minimum</span></span><br><span data-ttu-id="e884f-128">要件セット</span><span class="sxs-lookup"><span data-stu-id="e884f-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e884f-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="e884f-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="e884f-130">作成</span><span class="sxs-lookup"><span data-stu-id="e884f-130">Compose</span></span><br><span data-ttu-id="e884f-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="e884f-131">Read</span></span> | <span data-ttu-id="e884f-132">文字列</span><span class="sxs-lookup"><span data-stu-id="e884f-132">String</span></span> | [<span data-ttu-id="e884f-133">1.1</span><span class="sxs-lookup"><span data-stu-id="e884f-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e884f-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="e884f-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="e884f-135">作成</span><span class="sxs-lookup"><span data-stu-id="e884f-135">Compose</span></span><br><span data-ttu-id="e884f-136">読み取り</span><span class="sxs-lookup"><span data-stu-id="e884f-136">Read</span></span> | <span data-ttu-id="e884f-137">文字列</span><span class="sxs-lookup"><span data-stu-id="e884f-137">String</span></span> | [<span data-ttu-id="e884f-138">1.1</span><span class="sxs-lookup"><span data-stu-id="e884f-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e884f-139">EventType</span><span class="sxs-lookup"><span data-stu-id="e884f-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="e884f-140">作成</span><span class="sxs-lookup"><span data-stu-id="e884f-140">Compose</span></span><br><span data-ttu-id="e884f-141">読み取り</span><span class="sxs-lookup"><span data-stu-id="e884f-141">Read</span></span> | <span data-ttu-id="e884f-142">文字列</span><span class="sxs-lookup"><span data-stu-id="e884f-142">String</span></span> | [<span data-ttu-id="e884f-143">1.5</span><span class="sxs-lookup"><span data-stu-id="e884f-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="e884f-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="e884f-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="e884f-145">作成</span><span class="sxs-lookup"><span data-stu-id="e884f-145">Compose</span></span><br><span data-ttu-id="e884f-146">読み取り</span><span class="sxs-lookup"><span data-stu-id="e884f-146">Read</span></span> | <span data-ttu-id="e884f-147">文字列</span><span class="sxs-lookup"><span data-stu-id="e884f-147">String</span></span> | [<span data-ttu-id="e884f-148">1.1</span><span class="sxs-lookup"><span data-stu-id="e884f-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="e884f-149">名前空間</span><span class="sxs-lookup"><span data-stu-id="e884f-149">Namespaces</span></span>

<span data-ttu-id="e884f-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="e884f-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="e884f-151">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="e884f-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="e884f-152">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="e884f-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="e884f-153">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="e884f-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="e884f-154">型</span><span class="sxs-lookup"><span data-stu-id="e884f-154">Type</span></span>

*   <span data-ttu-id="e884f-155">String</span><span class="sxs-lookup"><span data-stu-id="e884f-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e884f-156">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="e884f-156">Properties:</span></span>

|<span data-ttu-id="e884f-157">名前</span><span class="sxs-lookup"><span data-stu-id="e884f-157">Name</span></span>| <span data-ttu-id="e884f-158">種類</span><span class="sxs-lookup"><span data-stu-id="e884f-158">Type</span></span>| <span data-ttu-id="e884f-159">説明</span><span class="sxs-lookup"><span data-stu-id="e884f-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="e884f-160">文字列</span><span class="sxs-lookup"><span data-stu-id="e884f-160">String</span></span>|<span data-ttu-id="e884f-161">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="e884f-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="e884f-162">文字列</span><span class="sxs-lookup"><span data-stu-id="e884f-162">String</span></span>|<span data-ttu-id="e884f-163">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="e884f-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e884f-164">Requirements</span><span class="sxs-lookup"><span data-stu-id="e884f-164">Requirements</span></span>

|<span data-ttu-id="e884f-165">要件</span><span class="sxs-lookup"><span data-stu-id="e884f-165">Requirement</span></span>| <span data-ttu-id="e884f-166">値</span><span class="sxs-lookup"><span data-stu-id="e884f-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="e884f-167">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e884f-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e884f-168">1.1</span><span class="sxs-lookup"><span data-stu-id="e884f-168">1.1</span></span>|
|[<span data-ttu-id="e884f-169">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e884f-169">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e884f-170">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e884f-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="e884f-171">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="e884f-171">CoercionType: String</span></span>

<span data-ttu-id="e884f-172">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="e884f-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e884f-173">型</span><span class="sxs-lookup"><span data-stu-id="e884f-173">Type</span></span>

*   <span data-ttu-id="e884f-174">String</span><span class="sxs-lookup"><span data-stu-id="e884f-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e884f-175">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="e884f-175">Properties:</span></span>

|<span data-ttu-id="e884f-176">名前</span><span class="sxs-lookup"><span data-stu-id="e884f-176">Name</span></span>| <span data-ttu-id="e884f-177">種類</span><span class="sxs-lookup"><span data-stu-id="e884f-177">Type</span></span>| <span data-ttu-id="e884f-178">説明</span><span class="sxs-lookup"><span data-stu-id="e884f-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="e884f-179">文字列</span><span class="sxs-lookup"><span data-stu-id="e884f-179">String</span></span>|<span data-ttu-id="e884f-180">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="e884f-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="e884f-181">String</span><span class="sxs-lookup"><span data-stu-id="e884f-181">String</span></span>|<span data-ttu-id="e884f-182">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="e884f-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e884f-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="e884f-183">Requirements</span></span>

|<span data-ttu-id="e884f-184">要件</span><span class="sxs-lookup"><span data-stu-id="e884f-184">Requirement</span></span>| <span data-ttu-id="e884f-185">値</span><span class="sxs-lookup"><span data-stu-id="e884f-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="e884f-186">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e884f-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e884f-187">1.1</span><span class="sxs-lookup"><span data-stu-id="e884f-187">1.1</span></span>|
|[<span data-ttu-id="e884f-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e884f-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e884f-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e884f-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="e884f-190">EventType: String</span><span class="sxs-lookup"><span data-stu-id="e884f-190">EventType: String</span></span>

<span data-ttu-id="e884f-191">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="e884f-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="e884f-192">型</span><span class="sxs-lookup"><span data-stu-id="e884f-192">Type</span></span>

*   <span data-ttu-id="e884f-193">String</span><span class="sxs-lookup"><span data-stu-id="e884f-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e884f-194">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="e884f-194">Properties:</span></span>

| <span data-ttu-id="e884f-195">名前</span><span class="sxs-lookup"><span data-stu-id="e884f-195">Name</span></span> | <span data-ttu-id="e884f-196">種類</span><span class="sxs-lookup"><span data-stu-id="e884f-196">Type</span></span> | <span data-ttu-id="e884f-197">説明</span><span class="sxs-lookup"><span data-stu-id="e884f-197">Description</span></span> | <span data-ttu-id="e884f-198">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="e884f-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="e884f-199">文字列</span><span class="sxs-lookup"><span data-stu-id="e884f-199">String</span></span> | <span data-ttu-id="e884f-200">選択した予定またはデータ系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="e884f-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="e884f-201">1.7</span><span class="sxs-lookup"><span data-stu-id="e884f-201">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="e884f-202">文字列</span><span class="sxs-lookup"><span data-stu-id="e884f-202">String</span></span> | <span data-ttu-id="e884f-203">添付ファイルがアイテムに追加またはアイテムから削除されています。</span><span class="sxs-lookup"><span data-stu-id="e884f-203">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="e884f-204">1.8</span><span class="sxs-lookup"><span data-stu-id="e884f-204">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="e884f-205">文字列</span><span class="sxs-lookup"><span data-stu-id="e884f-205">String</span></span> | <span data-ttu-id="e884f-206">選択した予定の場所が変更されました。</span><span class="sxs-lookup"><span data-stu-id="e884f-206">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="e884f-207">1.8</span><span class="sxs-lookup"><span data-stu-id="e884f-207">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="e884f-208">文字列</span><span class="sxs-lookup"><span data-stu-id="e884f-208">String</span></span> | <span data-ttu-id="e884f-209">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="e884f-209">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="e884f-210">1.5</span><span class="sxs-lookup"><span data-stu-id="e884f-210">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="e884f-211">文字列</span><span class="sxs-lookup"><span data-stu-id="e884f-211">String</span></span> | <span data-ttu-id="e884f-212">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="e884f-212">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="e884f-213">1.7</span><span class="sxs-lookup"><span data-stu-id="e884f-213">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="e884f-214">文字列</span><span class="sxs-lookup"><span data-stu-id="e884f-214">String</span></span> | <span data-ttu-id="e884f-215">選択したアイテムの定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="e884f-215">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="e884f-216">1.7</span><span class="sxs-lookup"><span data-stu-id="e884f-216">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e884f-217">Requirements</span><span class="sxs-lookup"><span data-stu-id="e884f-217">Requirements</span></span>

|<span data-ttu-id="e884f-218">要件</span><span class="sxs-lookup"><span data-stu-id="e884f-218">Requirement</span></span>| <span data-ttu-id="e884f-219">値</span><span class="sxs-lookup"><span data-stu-id="e884f-219">Value</span></span>|
|---|---|
|[<span data-ttu-id="e884f-220">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e884f-220">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e884f-221">1.5</span><span class="sxs-lookup"><span data-stu-id="e884f-221">1.5</span></span> |
|[<span data-ttu-id="e884f-222">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e884f-222">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e884f-223">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e884f-223">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="e884f-224">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="e884f-224">SourceProperty: String</span></span>

<span data-ttu-id="e884f-225">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="e884f-225">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e884f-226">型</span><span class="sxs-lookup"><span data-stu-id="e884f-226">Type</span></span>

*   <span data-ttu-id="e884f-227">String</span><span class="sxs-lookup"><span data-stu-id="e884f-227">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e884f-228">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="e884f-228">Properties:</span></span>

|<span data-ttu-id="e884f-229">名前</span><span class="sxs-lookup"><span data-stu-id="e884f-229">Name</span></span>| <span data-ttu-id="e884f-230">種類</span><span class="sxs-lookup"><span data-stu-id="e884f-230">Type</span></span>| <span data-ttu-id="e884f-231">説明</span><span class="sxs-lookup"><span data-stu-id="e884f-231">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="e884f-232">文字列</span><span class="sxs-lookup"><span data-stu-id="e884f-232">String</span></span>|<span data-ttu-id="e884f-233">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="e884f-233">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="e884f-234">文字列</span><span class="sxs-lookup"><span data-stu-id="e884f-234">String</span></span>|<span data-ttu-id="e884f-235">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="e884f-235">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e884f-236">Requirements</span><span class="sxs-lookup"><span data-stu-id="e884f-236">Requirements</span></span>

|<span data-ttu-id="e884f-237">要件</span><span class="sxs-lookup"><span data-stu-id="e884f-237">Requirement</span></span>| <span data-ttu-id="e884f-238">値</span><span class="sxs-lookup"><span data-stu-id="e884f-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="e884f-239">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e884f-239">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e884f-240">1.1</span><span class="sxs-lookup"><span data-stu-id="e884f-240">1.1</span></span>|
|[<span data-ttu-id="e884f-241">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e884f-241">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e884f-242">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e884f-242">Compose or Read</span></span>|
