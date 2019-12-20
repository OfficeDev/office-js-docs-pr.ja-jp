---
title: Office 名前空間-プレビュー要件セット
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: ef9634058fcdc633e9ad3a0adb74c4abebf8038b
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40815061"
---
# <a name="office"></a><span data-ttu-id="1df04-102">Office</span><span class="sxs-lookup"><span data-stu-id="1df04-102">Office</span></span>

<span data-ttu-id="1df04-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1df04-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="1df04-105">要件</span><span class="sxs-lookup"><span data-stu-id="1df04-105">Requirements</span></span>

|<span data-ttu-id="1df04-106">要件</span><span class="sxs-lookup"><span data-stu-id="1df04-106">Requirement</span></span>| <span data-ttu-id="1df04-107">値</span><span class="sxs-lookup"><span data-stu-id="1df04-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="1df04-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1df04-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1df04-109">1.1</span><span class="sxs-lookup"><span data-stu-id="1df04-109">1.1</span></span>|
|[<span data-ttu-id="1df04-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1df04-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1df04-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1df04-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="1df04-112">Properties</span><span class="sxs-lookup"><span data-stu-id="1df04-112">Properties</span></span>

| <span data-ttu-id="1df04-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="1df04-113">Property</span></span> | <span data-ttu-id="1df04-114">モード</span><span class="sxs-lookup"><span data-stu-id="1df04-114">Modes</span></span> | <span data-ttu-id="1df04-115">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="1df04-115">Return type</span></span> | <span data-ttu-id="1df04-116">最小値</span><span class="sxs-lookup"><span data-stu-id="1df04-116">Minimum</span></span><br><span data-ttu-id="1df04-117">要件セット</span><span class="sxs-lookup"><span data-stu-id="1df04-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="1df04-118">context</span><span class="sxs-lookup"><span data-stu-id="1df04-118">context</span></span>](office.context.md) | <span data-ttu-id="1df04-119">作成</span><span class="sxs-lookup"><span data-stu-id="1df04-119">Compose</span></span><br><span data-ttu-id="1df04-120">読み取り</span><span class="sxs-lookup"><span data-stu-id="1df04-120">Read</span></span> | [<span data-ttu-id="1df04-121">Context</span><span class="sxs-lookup"><span data-stu-id="1df04-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview) | [<span data-ttu-id="1df04-122">1.1</span><span class="sxs-lookup"><span data-stu-id="1df04-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="1df04-123">列挙型</span><span class="sxs-lookup"><span data-stu-id="1df04-123">Enumerations</span></span>

| <span data-ttu-id="1df04-124">列挙体</span><span class="sxs-lookup"><span data-stu-id="1df04-124">Enumeration</span></span> | <span data-ttu-id="1df04-125">モード</span><span class="sxs-lookup"><span data-stu-id="1df04-125">Modes</span></span> | <span data-ttu-id="1df04-126">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="1df04-126">Return type</span></span> | <span data-ttu-id="1df04-127">最小値</span><span class="sxs-lookup"><span data-stu-id="1df04-127">Minimum</span></span><br><span data-ttu-id="1df04-128">要件セット</span><span class="sxs-lookup"><span data-stu-id="1df04-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="1df04-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="1df04-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="1df04-130">作成</span><span class="sxs-lookup"><span data-stu-id="1df04-130">Compose</span></span><br><span data-ttu-id="1df04-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="1df04-131">Read</span></span> | <span data-ttu-id="1df04-132">String</span><span class="sxs-lookup"><span data-stu-id="1df04-132">String</span></span> | [<span data-ttu-id="1df04-133">1.1</span><span class="sxs-lookup"><span data-stu-id="1df04-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1df04-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="1df04-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="1df04-135">作成</span><span class="sxs-lookup"><span data-stu-id="1df04-135">Compose</span></span><br><span data-ttu-id="1df04-136">読み取り</span><span class="sxs-lookup"><span data-stu-id="1df04-136">Read</span></span> | <span data-ttu-id="1df04-137">String</span><span class="sxs-lookup"><span data-stu-id="1df04-137">String</span></span> | [<span data-ttu-id="1df04-138">1.1</span><span class="sxs-lookup"><span data-stu-id="1df04-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1df04-139">EventType</span><span class="sxs-lookup"><span data-stu-id="1df04-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="1df04-140">作成</span><span class="sxs-lookup"><span data-stu-id="1df04-140">Compose</span></span><br><span data-ttu-id="1df04-141">読み取り</span><span class="sxs-lookup"><span data-stu-id="1df04-141">Read</span></span> | <span data-ttu-id="1df04-142">String</span><span class="sxs-lookup"><span data-stu-id="1df04-142">String</span></span> | [<span data-ttu-id="1df04-143">1.5</span><span class="sxs-lookup"><span data-stu-id="1df04-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="1df04-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="1df04-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="1df04-145">作成</span><span class="sxs-lookup"><span data-stu-id="1df04-145">Compose</span></span><br><span data-ttu-id="1df04-146">読み取り</span><span class="sxs-lookup"><span data-stu-id="1df04-146">Read</span></span> | <span data-ttu-id="1df04-147">String</span><span class="sxs-lookup"><span data-stu-id="1df04-147">String</span></span> | [<span data-ttu-id="1df04-148">1.1</span><span class="sxs-lookup"><span data-stu-id="1df04-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="1df04-149">名前空間</span><span class="sxs-lookup"><span data-stu-id="1df04-149">Namespaces</span></span>

<span data-ttu-id="1df04-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="1df04-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="1df04-151">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="1df04-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="1df04-152">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="1df04-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="1df04-153">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="1df04-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="1df04-154">型</span><span class="sxs-lookup"><span data-stu-id="1df04-154">Type</span></span>

*   <span data-ttu-id="1df04-155">String</span><span class="sxs-lookup"><span data-stu-id="1df04-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1df04-156">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="1df04-156">Properties:</span></span>

|<span data-ttu-id="1df04-157">名前</span><span class="sxs-lookup"><span data-stu-id="1df04-157">Name</span></span>| <span data-ttu-id="1df04-158">種類</span><span class="sxs-lookup"><span data-stu-id="1df04-158">Type</span></span>| <span data-ttu-id="1df04-159">説明</span><span class="sxs-lookup"><span data-stu-id="1df04-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="1df04-160">String</span><span class="sxs-lookup"><span data-stu-id="1df04-160">String</span></span>|<span data-ttu-id="1df04-161">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="1df04-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="1df04-162">String</span><span class="sxs-lookup"><span data-stu-id="1df04-162">String</span></span>|<span data-ttu-id="1df04-163">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="1df04-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1df04-164">要件</span><span class="sxs-lookup"><span data-stu-id="1df04-164">Requirements</span></span>

|<span data-ttu-id="1df04-165">要件</span><span class="sxs-lookup"><span data-stu-id="1df04-165">Requirement</span></span>| <span data-ttu-id="1df04-166">値</span><span class="sxs-lookup"><span data-stu-id="1df04-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="1df04-167">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1df04-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1df04-168">1.1</span><span class="sxs-lookup"><span data-stu-id="1df04-168">1.1</span></span>|
|[<span data-ttu-id="1df04-169">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1df04-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1df04-170">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1df04-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="1df04-171">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="1df04-171">CoercionType: String</span></span>

<span data-ttu-id="1df04-172">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="1df04-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="1df04-173">型</span><span class="sxs-lookup"><span data-stu-id="1df04-173">Type</span></span>

*   <span data-ttu-id="1df04-174">String</span><span class="sxs-lookup"><span data-stu-id="1df04-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1df04-175">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="1df04-175">Properties:</span></span>

|<span data-ttu-id="1df04-176">名前</span><span class="sxs-lookup"><span data-stu-id="1df04-176">Name</span></span>| <span data-ttu-id="1df04-177">種類</span><span class="sxs-lookup"><span data-stu-id="1df04-177">Type</span></span>| <span data-ttu-id="1df04-178">説明</span><span class="sxs-lookup"><span data-stu-id="1df04-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="1df04-179">String</span><span class="sxs-lookup"><span data-stu-id="1df04-179">String</span></span>|<span data-ttu-id="1df04-180">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="1df04-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="1df04-181">String</span><span class="sxs-lookup"><span data-stu-id="1df04-181">String</span></span>|<span data-ttu-id="1df04-182">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="1df04-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1df04-183">要件</span><span class="sxs-lookup"><span data-stu-id="1df04-183">Requirements</span></span>

|<span data-ttu-id="1df04-184">要件</span><span class="sxs-lookup"><span data-stu-id="1df04-184">Requirement</span></span>| <span data-ttu-id="1df04-185">値</span><span class="sxs-lookup"><span data-stu-id="1df04-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="1df04-186">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1df04-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1df04-187">1.1</span><span class="sxs-lookup"><span data-stu-id="1df04-187">1.1</span></span>|
|[<span data-ttu-id="1df04-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1df04-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1df04-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1df04-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="1df04-190">EventType: String</span><span class="sxs-lookup"><span data-stu-id="1df04-190">EventType: String</span></span>

<span data-ttu-id="1df04-191">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="1df04-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="1df04-192">型</span><span class="sxs-lookup"><span data-stu-id="1df04-192">Type</span></span>

*   <span data-ttu-id="1df04-193">String</span><span class="sxs-lookup"><span data-stu-id="1df04-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1df04-194">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="1df04-194">Properties:</span></span>

| <span data-ttu-id="1df04-195">名前</span><span class="sxs-lookup"><span data-stu-id="1df04-195">Name</span></span> | <span data-ttu-id="1df04-196">種類</span><span class="sxs-lookup"><span data-stu-id="1df04-196">Type</span></span> | <span data-ttu-id="1df04-197">説明</span><span class="sxs-lookup"><span data-stu-id="1df04-197">Description</span></span> | <span data-ttu-id="1df04-198">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="1df04-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="1df04-199">String</span><span class="sxs-lookup"><span data-stu-id="1df04-199">String</span></span> | <span data-ttu-id="1df04-200">選択した予定またはデータ系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="1df04-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="1df04-201">1.7</span><span class="sxs-lookup"><span data-stu-id="1df04-201">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="1df04-202">String</span><span class="sxs-lookup"><span data-stu-id="1df04-202">String</span></span> | <span data-ttu-id="1df04-203">添付ファイルがアイテムに追加またはアイテムから削除されています。</span><span class="sxs-lookup"><span data-stu-id="1df04-203">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="1df04-204">1.8</span><span class="sxs-lookup"><span data-stu-id="1df04-204">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="1df04-205">String</span><span class="sxs-lookup"><span data-stu-id="1df04-205">String</span></span> | <span data-ttu-id="1df04-206">選択した予定の場所が変更されました。</span><span class="sxs-lookup"><span data-stu-id="1df04-206">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="1df04-207">1.8</span><span class="sxs-lookup"><span data-stu-id="1df04-207">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="1df04-208">String</span><span class="sxs-lookup"><span data-stu-id="1df04-208">String</span></span> | <span data-ttu-id="1df04-209">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="1df04-209">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="1df04-210">1.5</span><span class="sxs-lookup"><span data-stu-id="1df04-210">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="1df04-211">String</span><span class="sxs-lookup"><span data-stu-id="1df04-211">String</span></span> | <span data-ttu-id="1df04-212">メールボックスの Office テーマが変更されました。</span><span class="sxs-lookup"><span data-stu-id="1df04-212">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="1df04-213">プレビュー</span><span class="sxs-lookup"><span data-stu-id="1df04-213">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="1df04-214">String</span><span class="sxs-lookup"><span data-stu-id="1df04-214">String</span></span> | <span data-ttu-id="1df04-215">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="1df04-215">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="1df04-216">1.7</span><span class="sxs-lookup"><span data-stu-id="1df04-216">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="1df04-217">String</span><span class="sxs-lookup"><span data-stu-id="1df04-217">String</span></span> | <span data-ttu-id="1df04-218">選択したアイテムの定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="1df04-218">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="1df04-219">1.7</span><span class="sxs-lookup"><span data-stu-id="1df04-219">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1df04-220">要件</span><span class="sxs-lookup"><span data-stu-id="1df04-220">Requirements</span></span>

|<span data-ttu-id="1df04-221">要件</span><span class="sxs-lookup"><span data-stu-id="1df04-221">Requirement</span></span>| <span data-ttu-id="1df04-222">値</span><span class="sxs-lookup"><span data-stu-id="1df04-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="1df04-223">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1df04-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1df04-224">1.5</span><span class="sxs-lookup"><span data-stu-id="1df04-224">1.5</span></span> |
|[<span data-ttu-id="1df04-225">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1df04-225">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1df04-226">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1df04-226">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="1df04-227">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="1df04-227">SourceProperty: String</span></span>

<span data-ttu-id="1df04-228">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="1df04-228">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="1df04-229">型</span><span class="sxs-lookup"><span data-stu-id="1df04-229">Type</span></span>

*   <span data-ttu-id="1df04-230">String</span><span class="sxs-lookup"><span data-stu-id="1df04-230">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1df04-231">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="1df04-231">Properties:</span></span>

|<span data-ttu-id="1df04-232">名前</span><span class="sxs-lookup"><span data-stu-id="1df04-232">Name</span></span>| <span data-ttu-id="1df04-233">種類</span><span class="sxs-lookup"><span data-stu-id="1df04-233">Type</span></span>| <span data-ttu-id="1df04-234">説明</span><span class="sxs-lookup"><span data-stu-id="1df04-234">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="1df04-235">String</span><span class="sxs-lookup"><span data-stu-id="1df04-235">String</span></span>|<span data-ttu-id="1df04-236">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="1df04-236">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="1df04-237">String</span><span class="sxs-lookup"><span data-stu-id="1df04-237">String</span></span>|<span data-ttu-id="1df04-238">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="1df04-238">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1df04-239">要件</span><span class="sxs-lookup"><span data-stu-id="1df04-239">Requirements</span></span>

|<span data-ttu-id="1df04-240">要件</span><span class="sxs-lookup"><span data-stu-id="1df04-240">Requirement</span></span>| <span data-ttu-id="1df04-241">値</span><span class="sxs-lookup"><span data-stu-id="1df04-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="1df04-242">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1df04-242">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1df04-243">1.1</span><span class="sxs-lookup"><span data-stu-id="1df04-243">1.1</span></span>|
|[<span data-ttu-id="1df04-244">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1df04-244">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1df04-245">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1df04-245">Compose or Read</span></span>|
