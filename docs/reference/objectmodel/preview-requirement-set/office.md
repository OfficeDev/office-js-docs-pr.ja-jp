---
title: Office - プレビュー要件セット
description: Office API プレビュー要件セットを使用Outlookアドインで使用できる名前空間メンバーを指定します。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 72e2300dd50ff01e26417efaca92906049358fc0
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590884"
---
# <a name="office-mailbox-preview-requirement-set"></a><span data-ttu-id="61641-103">Office (メールボックス プレビュー要件セット)</span><span class="sxs-lookup"><span data-stu-id="61641-103">Office (Mailbox preview requirement set)</span></span>

<span data-ttu-id="61641-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="61641-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="61641-106">要件</span><span class="sxs-lookup"><span data-stu-id="61641-106">Requirements</span></span>

|<span data-ttu-id="61641-107">要件</span><span class="sxs-lookup"><span data-stu-id="61641-107">Requirement</span></span>| <span data-ttu-id="61641-108">値</span><span class="sxs-lookup"><span data-stu-id="61641-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="61641-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="61641-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="61641-110">1.1</span><span class="sxs-lookup"><span data-stu-id="61641-110">1.1</span></span>|
|[<span data-ttu-id="61641-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="61641-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="61641-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="61641-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="61641-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="61641-113">Properties</span></span>

| <span data-ttu-id="61641-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="61641-114">Property</span></span> | <span data-ttu-id="61641-115">モード</span><span class="sxs-lookup"><span data-stu-id="61641-115">Modes</span></span> | <span data-ttu-id="61641-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="61641-116">Return type</span></span> | <span data-ttu-id="61641-117">最小値</span><span class="sxs-lookup"><span data-stu-id="61641-117">Minimum</span></span><br><span data-ttu-id="61641-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="61641-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="61641-119">context</span><span class="sxs-lookup"><span data-stu-id="61641-119">context</span></span>](office.context.md) | <span data-ttu-id="61641-120">作成</span><span class="sxs-lookup"><span data-stu-id="61641-120">Compose</span></span><br><span data-ttu-id="61641-121">Read</span><span class="sxs-lookup"><span data-stu-id="61641-121">Read</span></span> | [<span data-ttu-id="61641-122">Context</span><span class="sxs-lookup"><span data-stu-id="61641-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="61641-123">1.1</span><span class="sxs-lookup"><span data-stu-id="61641-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="61641-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="61641-124">Enumerations</span></span>

| <span data-ttu-id="61641-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="61641-125">Enumeration</span></span> | <span data-ttu-id="61641-126">モード</span><span class="sxs-lookup"><span data-stu-id="61641-126">Modes</span></span> | <span data-ttu-id="61641-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="61641-127">Return type</span></span> | <span data-ttu-id="61641-128">最小値</span><span class="sxs-lookup"><span data-stu-id="61641-128">Minimum</span></span><br><span data-ttu-id="61641-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="61641-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="61641-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="61641-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="61641-131">作成</span><span class="sxs-lookup"><span data-stu-id="61641-131">Compose</span></span><br><span data-ttu-id="61641-132">Read</span><span class="sxs-lookup"><span data-stu-id="61641-132">Read</span></span> | <span data-ttu-id="61641-133">String</span><span class="sxs-lookup"><span data-stu-id="61641-133">String</span></span> | [<span data-ttu-id="61641-134">1.1</span><span class="sxs-lookup"><span data-stu-id="61641-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="61641-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="61641-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="61641-136">作成</span><span class="sxs-lookup"><span data-stu-id="61641-136">Compose</span></span><br><span data-ttu-id="61641-137">Read</span><span class="sxs-lookup"><span data-stu-id="61641-137">Read</span></span> | <span data-ttu-id="61641-138">String</span><span class="sxs-lookup"><span data-stu-id="61641-138">String</span></span> | [<span data-ttu-id="61641-139">1.1</span><span class="sxs-lookup"><span data-stu-id="61641-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="61641-140">EventType</span><span class="sxs-lookup"><span data-stu-id="61641-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="61641-141">作成</span><span class="sxs-lookup"><span data-stu-id="61641-141">Compose</span></span><br><span data-ttu-id="61641-142">Read</span><span class="sxs-lookup"><span data-stu-id="61641-142">Read</span></span> | <span data-ttu-id="61641-143">String</span><span class="sxs-lookup"><span data-stu-id="61641-143">String</span></span> | [<span data-ttu-id="61641-144">1.5</span><span class="sxs-lookup"><span data-stu-id="61641-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="61641-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="61641-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="61641-146">作成</span><span class="sxs-lookup"><span data-stu-id="61641-146">Compose</span></span><br><span data-ttu-id="61641-147">Read</span><span class="sxs-lookup"><span data-stu-id="61641-147">Read</span></span> | <span data-ttu-id="61641-148">String</span><span class="sxs-lookup"><span data-stu-id="61641-148">String</span></span> | [<span data-ttu-id="61641-149">1.1</span><span class="sxs-lookup"><span data-stu-id="61641-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="61641-150">名前空間</span><span class="sxs-lookup"><span data-stu-id="61641-150">Namespaces</span></span>

<span data-ttu-id="61641-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview&preserve-view=true): 、 など、Outlook固有の列挙の `ItemType` `EntityType` `AttachmentType` 数 `RecipientType` が `ResponseType` 含まれています `ItemNotificationMessageType` 。</span><span class="sxs-lookup"><span data-stu-id="61641-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="61641-152">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="61641-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="61641-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="61641-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="61641-154">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="61641-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="61641-155">型</span><span class="sxs-lookup"><span data-stu-id="61641-155">Type</span></span>

*   <span data-ttu-id="61641-156">String</span><span class="sxs-lookup"><span data-stu-id="61641-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="61641-157">プロパティ</span><span class="sxs-lookup"><span data-stu-id="61641-157">Properties</span></span>

|<span data-ttu-id="61641-158">名前</span><span class="sxs-lookup"><span data-stu-id="61641-158">Name</span></span>| <span data-ttu-id="61641-159">型</span><span class="sxs-lookup"><span data-stu-id="61641-159">Type</span></span>| <span data-ttu-id="61641-160">説明</span><span class="sxs-lookup"><span data-stu-id="61641-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="61641-161">String</span><span class="sxs-lookup"><span data-stu-id="61641-161">String</span></span>|<span data-ttu-id="61641-162">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="61641-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="61641-163">String</span><span class="sxs-lookup"><span data-stu-id="61641-163">String</span></span>|<span data-ttu-id="61641-164">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="61641-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="61641-165">要件</span><span class="sxs-lookup"><span data-stu-id="61641-165">Requirements</span></span>

|<span data-ttu-id="61641-166">要件</span><span class="sxs-lookup"><span data-stu-id="61641-166">Requirement</span></span>| <span data-ttu-id="61641-167">値</span><span class="sxs-lookup"><span data-stu-id="61641-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="61641-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="61641-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="61641-169">1.1</span><span class="sxs-lookup"><span data-stu-id="61641-169">1.1</span></span>|
|[<span data-ttu-id="61641-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="61641-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="61641-171">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="61641-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="61641-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="61641-172">CoercionType: String</span></span>

<span data-ttu-id="61641-173">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="61641-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="61641-174">型</span><span class="sxs-lookup"><span data-stu-id="61641-174">Type</span></span>

*   <span data-ttu-id="61641-175">String</span><span class="sxs-lookup"><span data-stu-id="61641-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="61641-176">プロパティ</span><span class="sxs-lookup"><span data-stu-id="61641-176">Properties</span></span>

|<span data-ttu-id="61641-177">名前</span><span class="sxs-lookup"><span data-stu-id="61641-177">Name</span></span>| <span data-ttu-id="61641-178">型</span><span class="sxs-lookup"><span data-stu-id="61641-178">Type</span></span>| <span data-ttu-id="61641-179">説明</span><span class="sxs-lookup"><span data-stu-id="61641-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="61641-180">String</span><span class="sxs-lookup"><span data-stu-id="61641-180">String</span></span>|<span data-ttu-id="61641-181">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="61641-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="61641-182">String</span><span class="sxs-lookup"><span data-stu-id="61641-182">String</span></span>|<span data-ttu-id="61641-183">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="61641-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="61641-184">要件</span><span class="sxs-lookup"><span data-stu-id="61641-184">Requirements</span></span>

|<span data-ttu-id="61641-185">要件</span><span class="sxs-lookup"><span data-stu-id="61641-185">Requirement</span></span>| <span data-ttu-id="61641-186">値</span><span class="sxs-lookup"><span data-stu-id="61641-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="61641-187">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="61641-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="61641-188">1.1</span><span class="sxs-lookup"><span data-stu-id="61641-188">1.1</span></span>|
|[<span data-ttu-id="61641-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="61641-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="61641-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="61641-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="61641-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="61641-191">EventType: String</span></span>

<span data-ttu-id="61641-192">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="61641-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="61641-193">型</span><span class="sxs-lookup"><span data-stu-id="61641-193">Type</span></span>

*   <span data-ttu-id="61641-194">String</span><span class="sxs-lookup"><span data-stu-id="61641-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="61641-195">プロパティ</span><span class="sxs-lookup"><span data-stu-id="61641-195">Properties</span></span>

| <span data-ttu-id="61641-196">名前</span><span class="sxs-lookup"><span data-stu-id="61641-196">Name</span></span> | <span data-ttu-id="61641-197">型</span><span class="sxs-lookup"><span data-stu-id="61641-197">Type</span></span> | <span data-ttu-id="61641-198">説明</span><span class="sxs-lookup"><span data-stu-id="61641-198">Description</span></span> | <span data-ttu-id="61641-199">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="61641-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="61641-200">String</span><span class="sxs-lookup"><span data-stu-id="61641-200">String</span></span> | <span data-ttu-id="61641-201">選択した予定または系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="61641-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="61641-202">1.7</span><span class="sxs-lookup"><span data-stu-id="61641-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="61641-203">String</span><span class="sxs-lookup"><span data-stu-id="61641-203">String</span></span> | <span data-ttu-id="61641-204">アイテムに添付ファイルが追加または削除されました。</span><span class="sxs-lookup"><span data-stu-id="61641-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="61641-205">1.8</span><span class="sxs-lookup"><span data-stu-id="61641-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="61641-206">String</span><span class="sxs-lookup"><span data-stu-id="61641-206">String</span></span> | <span data-ttu-id="61641-207">選択した予定の場所が変更されました。</span><span class="sxs-lookup"><span data-stu-id="61641-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="61641-208">1.8</span><span class="sxs-lookup"><span data-stu-id="61641-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="61641-209">String</span><span class="sxs-lookup"><span data-stu-id="61641-209">String</span></span> | <span data-ttu-id="61641-210">作業ウィンドウOutlook表示する場合は、別のアイテムが選択されています。</span><span class="sxs-lookup"><span data-stu-id="61641-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="61641-211">1.5</span><span class="sxs-lookup"><span data-stu-id="61641-211">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="61641-212">String</span><span class="sxs-lookup"><span data-stu-id="61641-212">String</span></span> | <span data-ttu-id="61641-213">メールボックスOfficeテーマが変更されました。</span><span class="sxs-lookup"><span data-stu-id="61641-213">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="61641-214">Preview</span><span class="sxs-lookup"><span data-stu-id="61641-214">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="61641-215">String</span><span class="sxs-lookup"><span data-stu-id="61641-215">String</span></span> | <span data-ttu-id="61641-216">選択したアイテムまたは予定の場所の受信者リストが変更されました。</span><span class="sxs-lookup"><span data-stu-id="61641-216">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="61641-217">1.7</span><span class="sxs-lookup"><span data-stu-id="61641-217">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="61641-218">String</span><span class="sxs-lookup"><span data-stu-id="61641-218">String</span></span> | <span data-ttu-id="61641-219">選択した系列の定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="61641-219">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="61641-220">1.7</span><span class="sxs-lookup"><span data-stu-id="61641-220">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="61641-221">要件</span><span class="sxs-lookup"><span data-stu-id="61641-221">Requirements</span></span>

|<span data-ttu-id="61641-222">要件</span><span class="sxs-lookup"><span data-stu-id="61641-222">Requirement</span></span>| <span data-ttu-id="61641-223">値</span><span class="sxs-lookup"><span data-stu-id="61641-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="61641-224">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="61641-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="61641-225">1.5</span><span class="sxs-lookup"><span data-stu-id="61641-225">1.5</span></span> |
|[<span data-ttu-id="61641-226">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="61641-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="61641-227">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="61641-227">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="61641-228">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="61641-228">SourceProperty: String</span></span>

<span data-ttu-id="61641-229">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="61641-229">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="61641-230">型</span><span class="sxs-lookup"><span data-stu-id="61641-230">Type</span></span>

*   <span data-ttu-id="61641-231">String</span><span class="sxs-lookup"><span data-stu-id="61641-231">String</span></span>

##### <a name="properties"></a><span data-ttu-id="61641-232">プロパティ</span><span class="sxs-lookup"><span data-stu-id="61641-232">Properties</span></span>

|<span data-ttu-id="61641-233">名前</span><span class="sxs-lookup"><span data-stu-id="61641-233">Name</span></span>| <span data-ttu-id="61641-234">型</span><span class="sxs-lookup"><span data-stu-id="61641-234">Type</span></span>| <span data-ttu-id="61641-235">説明</span><span class="sxs-lookup"><span data-stu-id="61641-235">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="61641-236">String</span><span class="sxs-lookup"><span data-stu-id="61641-236">String</span></span>|<span data-ttu-id="61641-237">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="61641-237">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="61641-238">String</span><span class="sxs-lookup"><span data-stu-id="61641-238">String</span></span>|<span data-ttu-id="61641-239">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="61641-239">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="61641-240">要件</span><span class="sxs-lookup"><span data-stu-id="61641-240">Requirements</span></span>

|<span data-ttu-id="61641-241">要件</span><span class="sxs-lookup"><span data-stu-id="61641-241">Requirement</span></span>| <span data-ttu-id="61641-242">値</span><span class="sxs-lookup"><span data-stu-id="61641-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="61641-243">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="61641-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="61641-244">1.1</span><span class="sxs-lookup"><span data-stu-id="61641-244">1.1</span></span>|
|[<span data-ttu-id="61641-245">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="61641-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="61641-246">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="61641-246">Compose or Read</span></span>|
