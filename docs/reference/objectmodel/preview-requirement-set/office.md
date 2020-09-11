---
title: Office 名前空間-プレビュー要件セット
description: メールボックス API プレビュー要件セットを使用した Outlook アドインで使用可能な Office 名前空間メンバー。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 1e0f932106df462c7cd172327082992f6e4d9a58
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431123"
---
# <a name="office-mailbox-preview-requirement-set"></a><span data-ttu-id="e0ddc-103">Office (メールボックスプレビュー要件セット)</span><span class="sxs-lookup"><span data-stu-id="e0ddc-103">Office (Mailbox preview requirement set)</span></span>

<span data-ttu-id="e0ddc-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e0ddc-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e0ddc-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="e0ddc-106">Requirements</span></span>

|<span data-ttu-id="e0ddc-107">要件</span><span class="sxs-lookup"><span data-stu-id="e0ddc-107">Requirement</span></span>| <span data-ttu-id="e0ddc-108">値</span><span class="sxs-lookup"><span data-stu-id="e0ddc-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="e0ddc-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e0ddc-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e0ddc-110">1.1</span><span class="sxs-lookup"><span data-stu-id="e0ddc-110">1.1</span></span>|
|[<span data-ttu-id="e0ddc-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e0ddc-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e0ddc-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e0ddc-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="e0ddc-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="e0ddc-113">Properties</span></span>

| <span data-ttu-id="e0ddc-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="e0ddc-114">Property</span></span> | <span data-ttu-id="e0ddc-115">モード</span><span class="sxs-lookup"><span data-stu-id="e0ddc-115">Modes</span></span> | <span data-ttu-id="e0ddc-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="e0ddc-116">Return type</span></span> | <span data-ttu-id="e0ddc-117">最小値</span><span class="sxs-lookup"><span data-stu-id="e0ddc-117">Minimum</span></span><br><span data-ttu-id="e0ddc-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="e0ddc-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e0ddc-119">context</span><span class="sxs-lookup"><span data-stu-id="e0ddc-119">context</span></span>](office.context.md) | <span data-ttu-id="e0ddc-120">作成</span><span class="sxs-lookup"><span data-stu-id="e0ddc-120">Compose</span></span><br><span data-ttu-id="e0ddc-121">読み取り</span><span class="sxs-lookup"><span data-stu-id="e0ddc-121">Read</span></span> | [<span data-ttu-id="e0ddc-122">Context</span><span class="sxs-lookup"><span data-stu-id="e0ddc-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="e0ddc-123">1.1</span><span class="sxs-lookup"><span data-stu-id="e0ddc-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="e0ddc-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="e0ddc-124">Enumerations</span></span>

| <span data-ttu-id="e0ddc-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="e0ddc-125">Enumeration</span></span> | <span data-ttu-id="e0ddc-126">モード</span><span class="sxs-lookup"><span data-stu-id="e0ddc-126">Modes</span></span> | <span data-ttu-id="e0ddc-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="e0ddc-127">Return type</span></span> | <span data-ttu-id="e0ddc-128">最小値</span><span class="sxs-lookup"><span data-stu-id="e0ddc-128">Minimum</span></span><br><span data-ttu-id="e0ddc-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="e0ddc-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e0ddc-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="e0ddc-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="e0ddc-131">作成</span><span class="sxs-lookup"><span data-stu-id="e0ddc-131">Compose</span></span><br><span data-ttu-id="e0ddc-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="e0ddc-132">Read</span></span> | <span data-ttu-id="e0ddc-133">文字列</span><span class="sxs-lookup"><span data-stu-id="e0ddc-133">String</span></span> | [<span data-ttu-id="e0ddc-134">1.1</span><span class="sxs-lookup"><span data-stu-id="e0ddc-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e0ddc-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="e0ddc-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="e0ddc-136">作成</span><span class="sxs-lookup"><span data-stu-id="e0ddc-136">Compose</span></span><br><span data-ttu-id="e0ddc-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="e0ddc-137">Read</span></span> | <span data-ttu-id="e0ddc-138">文字列</span><span class="sxs-lookup"><span data-stu-id="e0ddc-138">String</span></span> | [<span data-ttu-id="e0ddc-139">1.1</span><span class="sxs-lookup"><span data-stu-id="e0ddc-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e0ddc-140">EventType</span><span class="sxs-lookup"><span data-stu-id="e0ddc-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="e0ddc-141">作成</span><span class="sxs-lookup"><span data-stu-id="e0ddc-141">Compose</span></span><br><span data-ttu-id="e0ddc-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="e0ddc-142">Read</span></span> | <span data-ttu-id="e0ddc-143">文字列</span><span class="sxs-lookup"><span data-stu-id="e0ddc-143">String</span></span> | [<span data-ttu-id="e0ddc-144">1.5</span><span class="sxs-lookup"><span data-stu-id="e0ddc-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="e0ddc-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="e0ddc-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="e0ddc-146">作成</span><span class="sxs-lookup"><span data-stu-id="e0ddc-146">Compose</span></span><br><span data-ttu-id="e0ddc-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="e0ddc-147">Read</span></span> | <span data-ttu-id="e0ddc-148">文字列</span><span class="sxs-lookup"><span data-stu-id="e0ddc-148">String</span></span> | [<span data-ttu-id="e0ddc-149">1.1</span><span class="sxs-lookup"><span data-stu-id="e0ddc-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="e0ddc-150">名前空間</span><span class="sxs-lookup"><span data-stu-id="e0ddc-150">Namespaces</span></span>

<span data-ttu-id="e0ddc-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview&preserve-view=true):、、、、、など、多数の Outlook 固有の列挙を含み `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` ます。</span><span class="sxs-lookup"><span data-stu-id="e0ddc-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="e0ddc-152">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="e0ddc-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="e0ddc-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="e0ddc-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="e0ddc-154">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="e0ddc-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="e0ddc-155">型</span><span class="sxs-lookup"><span data-stu-id="e0ddc-155">Type</span></span>

*   <span data-ttu-id="e0ddc-156">String</span><span class="sxs-lookup"><span data-stu-id="e0ddc-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e0ddc-157">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="e0ddc-157">Properties:</span></span>

|<span data-ttu-id="e0ddc-158">名前</span><span class="sxs-lookup"><span data-stu-id="e0ddc-158">Name</span></span>| <span data-ttu-id="e0ddc-159">種類</span><span class="sxs-lookup"><span data-stu-id="e0ddc-159">Type</span></span>| <span data-ttu-id="e0ddc-160">説明</span><span class="sxs-lookup"><span data-stu-id="e0ddc-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="e0ddc-161">文字列</span><span class="sxs-lookup"><span data-stu-id="e0ddc-161">String</span></span>|<span data-ttu-id="e0ddc-162">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="e0ddc-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="e0ddc-163">String</span><span class="sxs-lookup"><span data-stu-id="e0ddc-163">String</span></span>|<span data-ttu-id="e0ddc-164">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="e0ddc-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e0ddc-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="e0ddc-165">Requirements</span></span>

|<span data-ttu-id="e0ddc-166">要件</span><span class="sxs-lookup"><span data-stu-id="e0ddc-166">Requirement</span></span>| <span data-ttu-id="e0ddc-167">値</span><span class="sxs-lookup"><span data-stu-id="e0ddc-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="e0ddc-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e0ddc-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e0ddc-169">1.1</span><span class="sxs-lookup"><span data-stu-id="e0ddc-169">1.1</span></span>|
|[<span data-ttu-id="e0ddc-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e0ddc-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e0ddc-171">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e0ddc-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="e0ddc-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="e0ddc-172">CoercionType: String</span></span>

<span data-ttu-id="e0ddc-173">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="e0ddc-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e0ddc-174">型</span><span class="sxs-lookup"><span data-stu-id="e0ddc-174">Type</span></span>

*   <span data-ttu-id="e0ddc-175">String</span><span class="sxs-lookup"><span data-stu-id="e0ddc-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e0ddc-176">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="e0ddc-176">Properties:</span></span>

|<span data-ttu-id="e0ddc-177">名前</span><span class="sxs-lookup"><span data-stu-id="e0ddc-177">Name</span></span>| <span data-ttu-id="e0ddc-178">種類</span><span class="sxs-lookup"><span data-stu-id="e0ddc-178">Type</span></span>| <span data-ttu-id="e0ddc-179">説明</span><span class="sxs-lookup"><span data-stu-id="e0ddc-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="e0ddc-180">文字列</span><span class="sxs-lookup"><span data-stu-id="e0ddc-180">String</span></span>|<span data-ttu-id="e0ddc-181">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="e0ddc-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="e0ddc-182">String</span><span class="sxs-lookup"><span data-stu-id="e0ddc-182">String</span></span>|<span data-ttu-id="e0ddc-183">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="e0ddc-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e0ddc-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="e0ddc-184">Requirements</span></span>

|<span data-ttu-id="e0ddc-185">要件</span><span class="sxs-lookup"><span data-stu-id="e0ddc-185">Requirement</span></span>| <span data-ttu-id="e0ddc-186">値</span><span class="sxs-lookup"><span data-stu-id="e0ddc-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="e0ddc-187">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e0ddc-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e0ddc-188">1.1</span><span class="sxs-lookup"><span data-stu-id="e0ddc-188">1.1</span></span>|
|[<span data-ttu-id="e0ddc-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e0ddc-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e0ddc-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e0ddc-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="e0ddc-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="e0ddc-191">EventType: String</span></span>

<span data-ttu-id="e0ddc-192">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="e0ddc-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="e0ddc-193">型</span><span class="sxs-lookup"><span data-stu-id="e0ddc-193">Type</span></span>

*   <span data-ttu-id="e0ddc-194">String</span><span class="sxs-lookup"><span data-stu-id="e0ddc-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e0ddc-195">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="e0ddc-195">Properties:</span></span>

| <span data-ttu-id="e0ddc-196">名前</span><span class="sxs-lookup"><span data-stu-id="e0ddc-196">Name</span></span> | <span data-ttu-id="e0ddc-197">種類</span><span class="sxs-lookup"><span data-stu-id="e0ddc-197">Type</span></span> | <span data-ttu-id="e0ddc-198">説明</span><span class="sxs-lookup"><span data-stu-id="e0ddc-198">Description</span></span> | <span data-ttu-id="e0ddc-199">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="e0ddc-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="e0ddc-200">文字列</span><span class="sxs-lookup"><span data-stu-id="e0ddc-200">String</span></span> | <span data-ttu-id="e0ddc-201">選択した予定またはデータ系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="e0ddc-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="e0ddc-202">1.7</span><span class="sxs-lookup"><span data-stu-id="e0ddc-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="e0ddc-203">文字列</span><span class="sxs-lookup"><span data-stu-id="e0ddc-203">String</span></span> | <span data-ttu-id="e0ddc-204">添付ファイルがアイテムに追加またはアイテムから削除されています。</span><span class="sxs-lookup"><span data-stu-id="e0ddc-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="e0ddc-205">1.8</span><span class="sxs-lookup"><span data-stu-id="e0ddc-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="e0ddc-206">文字列</span><span class="sxs-lookup"><span data-stu-id="e0ddc-206">String</span></span> | <span data-ttu-id="e0ddc-207">選択した予定の場所が変更されました。</span><span class="sxs-lookup"><span data-stu-id="e0ddc-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="e0ddc-208">1.8</span><span class="sxs-lookup"><span data-stu-id="e0ddc-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="e0ddc-209">文字列</span><span class="sxs-lookup"><span data-stu-id="e0ddc-209">String</span></span> | <span data-ttu-id="e0ddc-210">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="e0ddc-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="e0ddc-211">1.5</span><span class="sxs-lookup"><span data-stu-id="e0ddc-211">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="e0ddc-212">文字列</span><span class="sxs-lookup"><span data-stu-id="e0ddc-212">String</span></span> | <span data-ttu-id="e0ddc-213">メールボックスの Office テーマが変更されました。</span><span class="sxs-lookup"><span data-stu-id="e0ddc-213">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="e0ddc-214">Preview</span><span class="sxs-lookup"><span data-stu-id="e0ddc-214">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="e0ddc-215">文字列</span><span class="sxs-lookup"><span data-stu-id="e0ddc-215">String</span></span> | <span data-ttu-id="e0ddc-216">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="e0ddc-216">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="e0ddc-217">1.7</span><span class="sxs-lookup"><span data-stu-id="e0ddc-217">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="e0ddc-218">文字列</span><span class="sxs-lookup"><span data-stu-id="e0ddc-218">String</span></span> | <span data-ttu-id="e0ddc-219">選択したアイテムの定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="e0ddc-219">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="e0ddc-220">1.7</span><span class="sxs-lookup"><span data-stu-id="e0ddc-220">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e0ddc-221">Requirements</span><span class="sxs-lookup"><span data-stu-id="e0ddc-221">Requirements</span></span>

|<span data-ttu-id="e0ddc-222">要件</span><span class="sxs-lookup"><span data-stu-id="e0ddc-222">Requirement</span></span>| <span data-ttu-id="e0ddc-223">値</span><span class="sxs-lookup"><span data-stu-id="e0ddc-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="e0ddc-224">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e0ddc-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e0ddc-225">1.5</span><span class="sxs-lookup"><span data-stu-id="e0ddc-225">1.5</span></span> |
|[<span data-ttu-id="e0ddc-226">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e0ddc-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e0ddc-227">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e0ddc-227">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="e0ddc-228">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="e0ddc-228">SourceProperty: String</span></span>

<span data-ttu-id="e0ddc-229">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="e0ddc-229">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e0ddc-230">型</span><span class="sxs-lookup"><span data-stu-id="e0ddc-230">Type</span></span>

*   <span data-ttu-id="e0ddc-231">String</span><span class="sxs-lookup"><span data-stu-id="e0ddc-231">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e0ddc-232">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="e0ddc-232">Properties:</span></span>

|<span data-ttu-id="e0ddc-233">名前</span><span class="sxs-lookup"><span data-stu-id="e0ddc-233">Name</span></span>| <span data-ttu-id="e0ddc-234">種類</span><span class="sxs-lookup"><span data-stu-id="e0ddc-234">Type</span></span>| <span data-ttu-id="e0ddc-235">説明</span><span class="sxs-lookup"><span data-stu-id="e0ddc-235">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="e0ddc-236">文字列</span><span class="sxs-lookup"><span data-stu-id="e0ddc-236">String</span></span>|<span data-ttu-id="e0ddc-237">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="e0ddc-237">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="e0ddc-238">String</span><span class="sxs-lookup"><span data-stu-id="e0ddc-238">String</span></span>|<span data-ttu-id="e0ddc-239">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="e0ddc-239">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e0ddc-240">Requirements</span><span class="sxs-lookup"><span data-stu-id="e0ddc-240">Requirements</span></span>

|<span data-ttu-id="e0ddc-241">要件</span><span class="sxs-lookup"><span data-stu-id="e0ddc-241">Requirement</span></span>| <span data-ttu-id="e0ddc-242">値</span><span class="sxs-lookup"><span data-stu-id="e0ddc-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="e0ddc-243">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e0ddc-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e0ddc-244">1.1</span><span class="sxs-lookup"><span data-stu-id="e0ddc-244">1.1</span></span>|
|[<span data-ttu-id="e0ddc-245">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e0ddc-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e0ddc-246">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e0ddc-246">Compose or Read</span></span>|
