---
title: Office 名前空間-プレビュー要件セット
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: bd37b1be4d77d73cb56b0b2593ccc57dea6cab27
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629231"
---
# <a name="office"></a><span data-ttu-id="5b39b-102">Office</span><span class="sxs-lookup"><span data-stu-id="5b39b-102">Office</span></span>

<span data-ttu-id="5b39b-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5b39b-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="5b39b-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="5b39b-105">Requirements</span></span>

|<span data-ttu-id="5b39b-106">要件</span><span class="sxs-lookup"><span data-stu-id="5b39b-106">Requirement</span></span>| <span data-ttu-id="5b39b-107">値</span><span class="sxs-lookup"><span data-stu-id="5b39b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b39b-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5b39b-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b39b-109">1.0</span><span class="sxs-lookup"><span data-stu-id="5b39b-109">1.0</span></span>|
|[<span data-ttu-id="5b39b-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5b39b-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5b39b-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5b39b-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="5b39b-112">Properties</span><span class="sxs-lookup"><span data-stu-id="5b39b-112">Properties</span></span>

| <span data-ttu-id="5b39b-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="5b39b-113">Property</span></span> | <span data-ttu-id="5b39b-114">モード</span><span class="sxs-lookup"><span data-stu-id="5b39b-114">Modes</span></span> | <span data-ttu-id="5b39b-115">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="5b39b-115">Return type</span></span> | <span data-ttu-id="5b39b-116">最小値</span><span class="sxs-lookup"><span data-stu-id="5b39b-116">Minimum</span></span><br><span data-ttu-id="5b39b-117">要件セット</span><span class="sxs-lookup"><span data-stu-id="5b39b-117">requirement set</span></span> |
|---|---|---|---|
| [<span data-ttu-id="5b39b-118">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="5b39b-118">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="5b39b-119">作成</span><span class="sxs-lookup"><span data-stu-id="5b39b-119">Compose</span></span><br><span data-ttu-id="5b39b-120">読み取り</span><span class="sxs-lookup"><span data-stu-id="5b39b-120">Read</span></span> | <span data-ttu-id="5b39b-121">String</span><span class="sxs-lookup"><span data-stu-id="5b39b-121">String</span></span> | <span data-ttu-id="5b39b-122">1.0</span><span class="sxs-lookup"><span data-stu-id="5b39b-122">1.0</span></span> |
| [<span data-ttu-id="5b39b-123">CoercionType</span><span class="sxs-lookup"><span data-stu-id="5b39b-123">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="5b39b-124">作成</span><span class="sxs-lookup"><span data-stu-id="5b39b-124">Compose</span></span><br><span data-ttu-id="5b39b-125">読み取り</span><span class="sxs-lookup"><span data-stu-id="5b39b-125">Read</span></span> | <span data-ttu-id="5b39b-126">String</span><span class="sxs-lookup"><span data-stu-id="5b39b-126">String</span></span> | <span data-ttu-id="5b39b-127">1.0</span><span class="sxs-lookup"><span data-stu-id="5b39b-127">1.0</span></span> |
| [<span data-ttu-id="5b39b-128">EventType</span><span class="sxs-lookup"><span data-stu-id="5b39b-128">EventType</span></span>](#eventtype-string) | <span data-ttu-id="5b39b-129">作成</span><span class="sxs-lookup"><span data-stu-id="5b39b-129">Compose</span></span><br><span data-ttu-id="5b39b-130">読み取り</span><span class="sxs-lookup"><span data-stu-id="5b39b-130">Read</span></span> | <span data-ttu-id="5b39b-131">String</span><span class="sxs-lookup"><span data-stu-id="5b39b-131">String</span></span> | <span data-ttu-id="5b39b-132">1.5</span><span class="sxs-lookup"><span data-stu-id="5b39b-132">1.5</span></span> |
| [<span data-ttu-id="5b39b-133">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="5b39b-133">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="5b39b-134">作成</span><span class="sxs-lookup"><span data-stu-id="5b39b-134">Compose</span></span><br><span data-ttu-id="5b39b-135">読み取り</span><span class="sxs-lookup"><span data-stu-id="5b39b-135">Read</span></span> | <span data-ttu-id="5b39b-136">String</span><span class="sxs-lookup"><span data-stu-id="5b39b-136">String</span></span> | <span data-ttu-id="5b39b-137">1.0</span><span class="sxs-lookup"><span data-stu-id="5b39b-137">1.0</span></span> |

### <a name="namespaces"></a><span data-ttu-id="5b39b-138">名前空間</span><span class="sxs-lookup"><span data-stu-id="5b39b-138">Namespaces</span></span>

<span data-ttu-id="5b39b-139">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="5b39b-139">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="5b39b-140">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat):、、、、、、などのさまざま`ItemType`な`EntityType`列挙`AttachmentType` `RecipientType` `ResponseType`値が含まれ`ItemNotificationMessageType`ています。</span><span class="sxs-lookup"><span data-stu-id="5b39b-140">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="property-details"></a><span data-ttu-id="5b39b-141">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="5b39b-141">Property details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="5b39b-142">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="5b39b-142">AsyncResultStatus: String</span></span>

<span data-ttu-id="5b39b-143">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="5b39b-143">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="5b39b-144">型</span><span class="sxs-lookup"><span data-stu-id="5b39b-144">Type</span></span>

*   <span data-ttu-id="5b39b-145">String</span><span class="sxs-lookup"><span data-stu-id="5b39b-145">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5b39b-146">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="5b39b-146">Properties:</span></span>

|<span data-ttu-id="5b39b-147">名前</span><span class="sxs-lookup"><span data-stu-id="5b39b-147">Name</span></span>| <span data-ttu-id="5b39b-148">種類</span><span class="sxs-lookup"><span data-stu-id="5b39b-148">Type</span></span>| <span data-ttu-id="5b39b-149">説明</span><span class="sxs-lookup"><span data-stu-id="5b39b-149">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="5b39b-150">String</span><span class="sxs-lookup"><span data-stu-id="5b39b-150">String</span></span>|<span data-ttu-id="5b39b-151">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="5b39b-151">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="5b39b-152">String</span><span class="sxs-lookup"><span data-stu-id="5b39b-152">String</span></span>|<span data-ttu-id="5b39b-153">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="5b39b-153">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5b39b-154">Requirements</span><span class="sxs-lookup"><span data-stu-id="5b39b-154">Requirements</span></span>

|<span data-ttu-id="5b39b-155">要件</span><span class="sxs-lookup"><span data-stu-id="5b39b-155">Requirement</span></span>| <span data-ttu-id="5b39b-156">値</span><span class="sxs-lookup"><span data-stu-id="5b39b-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b39b-157">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5b39b-157">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b39b-158">1.0</span><span class="sxs-lookup"><span data-stu-id="5b39b-158">1.0</span></span>|
|[<span data-ttu-id="5b39b-159">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5b39b-159">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5b39b-160">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5b39b-160">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="5b39b-161">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="5b39b-161">CoercionType: String</span></span>

<span data-ttu-id="5b39b-162">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="5b39b-162">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5b39b-163">型</span><span class="sxs-lookup"><span data-stu-id="5b39b-163">Type</span></span>

*   <span data-ttu-id="5b39b-164">String</span><span class="sxs-lookup"><span data-stu-id="5b39b-164">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5b39b-165">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="5b39b-165">Properties:</span></span>

|<span data-ttu-id="5b39b-166">名前</span><span class="sxs-lookup"><span data-stu-id="5b39b-166">Name</span></span>| <span data-ttu-id="5b39b-167">種類</span><span class="sxs-lookup"><span data-stu-id="5b39b-167">Type</span></span>| <span data-ttu-id="5b39b-168">説明</span><span class="sxs-lookup"><span data-stu-id="5b39b-168">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="5b39b-169">String</span><span class="sxs-lookup"><span data-stu-id="5b39b-169">String</span></span>|<span data-ttu-id="5b39b-170">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="5b39b-170">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="5b39b-171">String</span><span class="sxs-lookup"><span data-stu-id="5b39b-171">String</span></span>|<span data-ttu-id="5b39b-172">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="5b39b-172">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5b39b-173">Requirements</span><span class="sxs-lookup"><span data-stu-id="5b39b-173">Requirements</span></span>

|<span data-ttu-id="5b39b-174">要件</span><span class="sxs-lookup"><span data-stu-id="5b39b-174">Requirement</span></span>| <span data-ttu-id="5b39b-175">値</span><span class="sxs-lookup"><span data-stu-id="5b39b-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b39b-176">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5b39b-176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b39b-177">1.0</span><span class="sxs-lookup"><span data-stu-id="5b39b-177">1.0</span></span>|
|[<span data-ttu-id="5b39b-178">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5b39b-178">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5b39b-179">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5b39b-179">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="5b39b-180">EventType: String</span><span class="sxs-lookup"><span data-stu-id="5b39b-180">EventType: String</span></span>

<span data-ttu-id="5b39b-181">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="5b39b-181">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="5b39b-182">型</span><span class="sxs-lookup"><span data-stu-id="5b39b-182">Type</span></span>

*   <span data-ttu-id="5b39b-183">String</span><span class="sxs-lookup"><span data-stu-id="5b39b-183">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5b39b-184">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="5b39b-184">Properties:</span></span>

| <span data-ttu-id="5b39b-185">名前</span><span class="sxs-lookup"><span data-stu-id="5b39b-185">Name</span></span> | <span data-ttu-id="5b39b-186">種類</span><span class="sxs-lookup"><span data-stu-id="5b39b-186">Type</span></span> | <span data-ttu-id="5b39b-187">説明</span><span class="sxs-lookup"><span data-stu-id="5b39b-187">Description</span></span> | <span data-ttu-id="5b39b-188">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="5b39b-188">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="5b39b-189">String</span><span class="sxs-lookup"><span data-stu-id="5b39b-189">String</span></span> | <span data-ttu-id="5b39b-190">選択した予定またはデータ系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="5b39b-190">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="5b39b-191">1.7</span><span class="sxs-lookup"><span data-stu-id="5b39b-191">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="5b39b-192">String</span><span class="sxs-lookup"><span data-stu-id="5b39b-192">String</span></span> | <span data-ttu-id="5b39b-193">添付ファイルがアイテムに追加またはアイテムから削除されています。</span><span class="sxs-lookup"><span data-stu-id="5b39b-193">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="5b39b-194">1.8</span><span class="sxs-lookup"><span data-stu-id="5b39b-194">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="5b39b-195">String</span><span class="sxs-lookup"><span data-stu-id="5b39b-195">String</span></span> | <span data-ttu-id="5b39b-196">選択した予定の場所が変更されました。</span><span class="sxs-lookup"><span data-stu-id="5b39b-196">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="5b39b-197">1.8</span><span class="sxs-lookup"><span data-stu-id="5b39b-197">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="5b39b-198">String</span><span class="sxs-lookup"><span data-stu-id="5b39b-198">String</span></span> | <span data-ttu-id="5b39b-199">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="5b39b-199">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="5b39b-200">1.5</span><span class="sxs-lookup"><span data-stu-id="5b39b-200">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="5b39b-201">String</span><span class="sxs-lookup"><span data-stu-id="5b39b-201">String</span></span> | <span data-ttu-id="5b39b-202">メールボックスの Office テーマが変更されました。</span><span class="sxs-lookup"><span data-stu-id="5b39b-202">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="5b39b-203">プレビュー</span><span class="sxs-lookup"><span data-stu-id="5b39b-203">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="5b39b-204">String</span><span class="sxs-lookup"><span data-stu-id="5b39b-204">String</span></span> | <span data-ttu-id="5b39b-205">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="5b39b-205">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="5b39b-206">1.7</span><span class="sxs-lookup"><span data-stu-id="5b39b-206">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="5b39b-207">String</span><span class="sxs-lookup"><span data-stu-id="5b39b-207">String</span></span> | <span data-ttu-id="5b39b-208">選択したアイテムの定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="5b39b-208">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="5b39b-209">1.7</span><span class="sxs-lookup"><span data-stu-id="5b39b-209">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5b39b-210">Requirements</span><span class="sxs-lookup"><span data-stu-id="5b39b-210">Requirements</span></span>

|<span data-ttu-id="5b39b-211">要件</span><span class="sxs-lookup"><span data-stu-id="5b39b-211">Requirement</span></span>| <span data-ttu-id="5b39b-212">値</span><span class="sxs-lookup"><span data-stu-id="5b39b-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b39b-213">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5b39b-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b39b-214">1.5</span><span class="sxs-lookup"><span data-stu-id="5b39b-214">1.5</span></span> |
|[<span data-ttu-id="5b39b-215">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5b39b-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5b39b-216">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5b39b-216">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="5b39b-217">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="5b39b-217">SourceProperty: String</span></span>

<span data-ttu-id="5b39b-218">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="5b39b-218">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5b39b-219">型</span><span class="sxs-lookup"><span data-stu-id="5b39b-219">Type</span></span>

*   <span data-ttu-id="5b39b-220">String</span><span class="sxs-lookup"><span data-stu-id="5b39b-220">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5b39b-221">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="5b39b-221">Properties:</span></span>

|<span data-ttu-id="5b39b-222">名前</span><span class="sxs-lookup"><span data-stu-id="5b39b-222">Name</span></span>| <span data-ttu-id="5b39b-223">種類</span><span class="sxs-lookup"><span data-stu-id="5b39b-223">Type</span></span>| <span data-ttu-id="5b39b-224">説明</span><span class="sxs-lookup"><span data-stu-id="5b39b-224">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="5b39b-225">String</span><span class="sxs-lookup"><span data-stu-id="5b39b-225">String</span></span>|<span data-ttu-id="5b39b-226">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="5b39b-226">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="5b39b-227">String</span><span class="sxs-lookup"><span data-stu-id="5b39b-227">String</span></span>|<span data-ttu-id="5b39b-228">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="5b39b-228">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5b39b-229">Requirements</span><span class="sxs-lookup"><span data-stu-id="5b39b-229">Requirements</span></span>

|<span data-ttu-id="5b39b-230">要件</span><span class="sxs-lookup"><span data-stu-id="5b39b-230">Requirement</span></span>| <span data-ttu-id="5b39b-231">値</span><span class="sxs-lookup"><span data-stu-id="5b39b-231">Value</span></span>|
|---|---|
|[<span data-ttu-id="5b39b-232">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5b39b-232">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5b39b-233">1.0</span><span class="sxs-lookup"><span data-stu-id="5b39b-233">1.0</span></span>|
|[<span data-ttu-id="5b39b-234">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5b39b-234">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5b39b-235">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5b39b-235">Compose or Read</span></span>|
