---
title: Office 名前空間-要件セット1.8
description: メールボックス API 要件セット1.8 を使用した Outlook アドインで使用可能な Office 名前空間メンバー。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: e0580cd1bb327c8673c46d3d0292aec9f2f1c971
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431522"
---
# <a name="office-mailbox-requirement-set-18"></a><span data-ttu-id="dc3d1-103">Office (メールボックス要件セット 1.8)</span><span class="sxs-lookup"><span data-stu-id="dc3d1-103">Office (Mailbox requirement set 1.8)</span></span>

<span data-ttu-id="dc3d1-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="dc3d1-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="dc3d1-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="dc3d1-106">Requirements</span></span>

|<span data-ttu-id="dc3d1-107">要件</span><span class="sxs-lookup"><span data-stu-id="dc3d1-107">Requirement</span></span>| <span data-ttu-id="dc3d1-108">値</span><span class="sxs-lookup"><span data-stu-id="dc3d1-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc3d1-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dc3d1-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dc3d1-110">1.1</span><span class="sxs-lookup"><span data-stu-id="dc3d1-110">1.1</span></span>|
|[<span data-ttu-id="dc3d1-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dc3d1-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dc3d1-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dc3d1-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="dc3d1-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="dc3d1-113">Properties</span></span>

| <span data-ttu-id="dc3d1-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="dc3d1-114">Property</span></span> | <span data-ttu-id="dc3d1-115">モード</span><span class="sxs-lookup"><span data-stu-id="dc3d1-115">Modes</span></span> | <span data-ttu-id="dc3d1-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="dc3d1-116">Return type</span></span> | <span data-ttu-id="dc3d1-117">最小値</span><span class="sxs-lookup"><span data-stu-id="dc3d1-117">Minimum</span></span><br><span data-ttu-id="dc3d1-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="dc3d1-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="dc3d1-119">context</span><span class="sxs-lookup"><span data-stu-id="dc3d1-119">context</span></span>](office.context.md) | <span data-ttu-id="dc3d1-120">作成</span><span class="sxs-lookup"><span data-stu-id="dc3d1-120">Compose</span></span><br><span data-ttu-id="dc3d1-121">読み取り</span><span class="sxs-lookup"><span data-stu-id="dc3d1-121">Read</span></span> | [<span data-ttu-id="dc3d1-122">Context</span><span class="sxs-lookup"><span data-stu-id="dc3d1-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="dc3d1-123">1.1</span><span class="sxs-lookup"><span data-stu-id="dc3d1-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="dc3d1-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="dc3d1-124">Enumerations</span></span>

| <span data-ttu-id="dc3d1-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="dc3d1-125">Enumeration</span></span> | <span data-ttu-id="dc3d1-126">モード</span><span class="sxs-lookup"><span data-stu-id="dc3d1-126">Modes</span></span> | <span data-ttu-id="dc3d1-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="dc3d1-127">Return type</span></span> | <span data-ttu-id="dc3d1-128">最小値</span><span class="sxs-lookup"><span data-stu-id="dc3d1-128">Minimum</span></span><br><span data-ttu-id="dc3d1-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="dc3d1-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="dc3d1-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="dc3d1-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="dc3d1-131">作成</span><span class="sxs-lookup"><span data-stu-id="dc3d1-131">Compose</span></span><br><span data-ttu-id="dc3d1-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="dc3d1-132">Read</span></span> | <span data-ttu-id="dc3d1-133">文字列</span><span class="sxs-lookup"><span data-stu-id="dc3d1-133">String</span></span> | [<span data-ttu-id="dc3d1-134">1.1</span><span class="sxs-lookup"><span data-stu-id="dc3d1-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dc3d1-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="dc3d1-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="dc3d1-136">作成</span><span class="sxs-lookup"><span data-stu-id="dc3d1-136">Compose</span></span><br><span data-ttu-id="dc3d1-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="dc3d1-137">Read</span></span> | <span data-ttu-id="dc3d1-138">文字列</span><span class="sxs-lookup"><span data-stu-id="dc3d1-138">String</span></span> | [<span data-ttu-id="dc3d1-139">1.1</span><span class="sxs-lookup"><span data-stu-id="dc3d1-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dc3d1-140">EventType</span><span class="sxs-lookup"><span data-stu-id="dc3d1-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="dc3d1-141">作成</span><span class="sxs-lookup"><span data-stu-id="dc3d1-141">Compose</span></span><br><span data-ttu-id="dc3d1-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="dc3d1-142">Read</span></span> | <span data-ttu-id="dc3d1-143">文字列</span><span class="sxs-lookup"><span data-stu-id="dc3d1-143">String</span></span> | [<span data-ttu-id="dc3d1-144">1.5</span><span class="sxs-lookup"><span data-stu-id="dc3d1-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="dc3d1-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="dc3d1-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="dc3d1-146">作成</span><span class="sxs-lookup"><span data-stu-id="dc3d1-146">Compose</span></span><br><span data-ttu-id="dc3d1-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="dc3d1-147">Read</span></span> | <span data-ttu-id="dc3d1-148">文字列</span><span class="sxs-lookup"><span data-stu-id="dc3d1-148">String</span></span> | [<span data-ttu-id="dc3d1-149">1.1</span><span class="sxs-lookup"><span data-stu-id="dc3d1-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="dc3d1-150">名前空間</span><span class="sxs-lookup"><span data-stu-id="dc3d1-150">Namespaces</span></span>

<span data-ttu-id="dc3d1-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8&preserve-view=true):、、、、、など、多数の Outlook 固有の列挙を含み `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` ます。</span><span class="sxs-lookup"><span data-stu-id="dc3d1-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="dc3d1-152">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="dc3d1-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="dc3d1-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="dc3d1-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="dc3d1-154">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="dc3d1-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="dc3d1-155">型</span><span class="sxs-lookup"><span data-stu-id="dc3d1-155">Type</span></span>

*   <span data-ttu-id="dc3d1-156">String</span><span class="sxs-lookup"><span data-stu-id="dc3d1-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dc3d1-157">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="dc3d1-157">Properties:</span></span>

|<span data-ttu-id="dc3d1-158">名前</span><span class="sxs-lookup"><span data-stu-id="dc3d1-158">Name</span></span>| <span data-ttu-id="dc3d1-159">種類</span><span class="sxs-lookup"><span data-stu-id="dc3d1-159">Type</span></span>| <span data-ttu-id="dc3d1-160">説明</span><span class="sxs-lookup"><span data-stu-id="dc3d1-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="dc3d1-161">文字列</span><span class="sxs-lookup"><span data-stu-id="dc3d1-161">String</span></span>|<span data-ttu-id="dc3d1-162">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="dc3d1-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="dc3d1-163">String</span><span class="sxs-lookup"><span data-stu-id="dc3d1-163">String</span></span>|<span data-ttu-id="dc3d1-164">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="dc3d1-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc3d1-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="dc3d1-165">Requirements</span></span>

|<span data-ttu-id="dc3d1-166">要件</span><span class="sxs-lookup"><span data-stu-id="dc3d1-166">Requirement</span></span>| <span data-ttu-id="dc3d1-167">値</span><span class="sxs-lookup"><span data-stu-id="dc3d1-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc3d1-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dc3d1-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dc3d1-169">1.1</span><span class="sxs-lookup"><span data-stu-id="dc3d1-169">1.1</span></span>|
|[<span data-ttu-id="dc3d1-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dc3d1-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dc3d1-171">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dc3d1-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="dc3d1-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="dc3d1-172">CoercionType: String</span></span>

<span data-ttu-id="dc3d1-173">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="dc3d1-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="dc3d1-174">型</span><span class="sxs-lookup"><span data-stu-id="dc3d1-174">Type</span></span>

*   <span data-ttu-id="dc3d1-175">String</span><span class="sxs-lookup"><span data-stu-id="dc3d1-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dc3d1-176">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="dc3d1-176">Properties:</span></span>

|<span data-ttu-id="dc3d1-177">名前</span><span class="sxs-lookup"><span data-stu-id="dc3d1-177">Name</span></span>| <span data-ttu-id="dc3d1-178">種類</span><span class="sxs-lookup"><span data-stu-id="dc3d1-178">Type</span></span>| <span data-ttu-id="dc3d1-179">説明</span><span class="sxs-lookup"><span data-stu-id="dc3d1-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="dc3d1-180">文字列</span><span class="sxs-lookup"><span data-stu-id="dc3d1-180">String</span></span>|<span data-ttu-id="dc3d1-181">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="dc3d1-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="dc3d1-182">String</span><span class="sxs-lookup"><span data-stu-id="dc3d1-182">String</span></span>|<span data-ttu-id="dc3d1-183">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="dc3d1-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc3d1-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="dc3d1-184">Requirements</span></span>

|<span data-ttu-id="dc3d1-185">要件</span><span class="sxs-lookup"><span data-stu-id="dc3d1-185">Requirement</span></span>| <span data-ttu-id="dc3d1-186">値</span><span class="sxs-lookup"><span data-stu-id="dc3d1-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc3d1-187">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dc3d1-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dc3d1-188">1.1</span><span class="sxs-lookup"><span data-stu-id="dc3d1-188">1.1</span></span>|
|[<span data-ttu-id="dc3d1-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dc3d1-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dc3d1-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dc3d1-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="dc3d1-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="dc3d1-191">EventType: String</span></span>

<span data-ttu-id="dc3d1-192">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="dc3d1-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="dc3d1-193">型</span><span class="sxs-lookup"><span data-stu-id="dc3d1-193">Type</span></span>

*   <span data-ttu-id="dc3d1-194">String</span><span class="sxs-lookup"><span data-stu-id="dc3d1-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dc3d1-195">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="dc3d1-195">Properties:</span></span>

| <span data-ttu-id="dc3d1-196">名前</span><span class="sxs-lookup"><span data-stu-id="dc3d1-196">Name</span></span> | <span data-ttu-id="dc3d1-197">種類</span><span class="sxs-lookup"><span data-stu-id="dc3d1-197">Type</span></span> | <span data-ttu-id="dc3d1-198">説明</span><span class="sxs-lookup"><span data-stu-id="dc3d1-198">Description</span></span> | <span data-ttu-id="dc3d1-199">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="dc3d1-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="dc3d1-200">文字列</span><span class="sxs-lookup"><span data-stu-id="dc3d1-200">String</span></span> | <span data-ttu-id="dc3d1-201">選択した予定またはデータ系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="dc3d1-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="dc3d1-202">1.7</span><span class="sxs-lookup"><span data-stu-id="dc3d1-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="dc3d1-203">文字列</span><span class="sxs-lookup"><span data-stu-id="dc3d1-203">String</span></span> | <span data-ttu-id="dc3d1-204">添付ファイルがアイテムに追加またはアイテムから削除されています。</span><span class="sxs-lookup"><span data-stu-id="dc3d1-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="dc3d1-205">1.8</span><span class="sxs-lookup"><span data-stu-id="dc3d1-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="dc3d1-206">文字列</span><span class="sxs-lookup"><span data-stu-id="dc3d1-206">String</span></span> | <span data-ttu-id="dc3d1-207">選択した予定の場所が変更されました。</span><span class="sxs-lookup"><span data-stu-id="dc3d1-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="dc3d1-208">1.8</span><span class="sxs-lookup"><span data-stu-id="dc3d1-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="dc3d1-209">文字列</span><span class="sxs-lookup"><span data-stu-id="dc3d1-209">String</span></span> | <span data-ttu-id="dc3d1-210">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="dc3d1-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="dc3d1-211">1.5</span><span class="sxs-lookup"><span data-stu-id="dc3d1-211">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="dc3d1-212">文字列</span><span class="sxs-lookup"><span data-stu-id="dc3d1-212">String</span></span> | <span data-ttu-id="dc3d1-213">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="dc3d1-213">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="dc3d1-214">1.7</span><span class="sxs-lookup"><span data-stu-id="dc3d1-214">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="dc3d1-215">文字列</span><span class="sxs-lookup"><span data-stu-id="dc3d1-215">String</span></span> | <span data-ttu-id="dc3d1-216">選択したアイテムの定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="dc3d1-216">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="dc3d1-217">1.7</span><span class="sxs-lookup"><span data-stu-id="dc3d1-217">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dc3d1-218">Requirements</span><span class="sxs-lookup"><span data-stu-id="dc3d1-218">Requirements</span></span>

|<span data-ttu-id="dc3d1-219">要件</span><span class="sxs-lookup"><span data-stu-id="dc3d1-219">Requirement</span></span>| <span data-ttu-id="dc3d1-220">値</span><span class="sxs-lookup"><span data-stu-id="dc3d1-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc3d1-221">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dc3d1-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dc3d1-222">1.5</span><span class="sxs-lookup"><span data-stu-id="dc3d1-222">1.5</span></span> |
|[<span data-ttu-id="dc3d1-223">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dc3d1-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dc3d1-224">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dc3d1-224">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="dc3d1-225">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="dc3d1-225">SourceProperty: String</span></span>

<span data-ttu-id="dc3d1-226">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="dc3d1-226">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="dc3d1-227">型</span><span class="sxs-lookup"><span data-stu-id="dc3d1-227">Type</span></span>

*   <span data-ttu-id="dc3d1-228">String</span><span class="sxs-lookup"><span data-stu-id="dc3d1-228">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dc3d1-229">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="dc3d1-229">Properties:</span></span>

|<span data-ttu-id="dc3d1-230">名前</span><span class="sxs-lookup"><span data-stu-id="dc3d1-230">Name</span></span>| <span data-ttu-id="dc3d1-231">種類</span><span class="sxs-lookup"><span data-stu-id="dc3d1-231">Type</span></span>| <span data-ttu-id="dc3d1-232">説明</span><span class="sxs-lookup"><span data-stu-id="dc3d1-232">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="dc3d1-233">文字列</span><span class="sxs-lookup"><span data-stu-id="dc3d1-233">String</span></span>|<span data-ttu-id="dc3d1-234">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="dc3d1-234">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="dc3d1-235">String</span><span class="sxs-lookup"><span data-stu-id="dc3d1-235">String</span></span>|<span data-ttu-id="dc3d1-236">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="dc3d1-236">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc3d1-237">Requirements</span><span class="sxs-lookup"><span data-stu-id="dc3d1-237">Requirements</span></span>

|<span data-ttu-id="dc3d1-238">要件</span><span class="sxs-lookup"><span data-stu-id="dc3d1-238">Requirement</span></span>| <span data-ttu-id="dc3d1-239">値</span><span class="sxs-lookup"><span data-stu-id="dc3d1-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc3d1-240">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dc3d1-240">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dc3d1-241">1.1</span><span class="sxs-lookup"><span data-stu-id="dc3d1-241">1.1</span></span>|
|[<span data-ttu-id="dc3d1-242">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dc3d1-242">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dc3d1-243">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dc3d1-243">Compose or Read</span></span>|
