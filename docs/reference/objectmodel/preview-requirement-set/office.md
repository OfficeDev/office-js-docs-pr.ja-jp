---
title: Office 名前空間-プレビュー要件セット
description: Outlook アドイン API の最上位レベルの名前空間のオブジェクトモデル (メールボックス API プレビューバージョン)。
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 40623c02fae820926d9162903320f30e5a424544
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720275"
---
# <a name="office"></a><span data-ttu-id="becf2-103">Office</span><span class="sxs-lookup"><span data-stu-id="becf2-103">Office</span></span>

<span data-ttu-id="becf2-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="becf2-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="becf2-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="becf2-106">Requirements</span></span>

|<span data-ttu-id="becf2-107">要件</span><span class="sxs-lookup"><span data-stu-id="becf2-107">Requirement</span></span>| <span data-ttu-id="becf2-108">値</span><span class="sxs-lookup"><span data-stu-id="becf2-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="becf2-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="becf2-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="becf2-110">1.1</span><span class="sxs-lookup"><span data-stu-id="becf2-110">1.1</span></span>|
|[<span data-ttu-id="becf2-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="becf2-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="becf2-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="becf2-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="becf2-113">Properties</span><span class="sxs-lookup"><span data-stu-id="becf2-113">Properties</span></span>

| <span data-ttu-id="becf2-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="becf2-114">Property</span></span> | <span data-ttu-id="becf2-115">モード</span><span class="sxs-lookup"><span data-stu-id="becf2-115">Modes</span></span> | <span data-ttu-id="becf2-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="becf2-116">Return type</span></span> | <span data-ttu-id="becf2-117">最小値</span><span class="sxs-lookup"><span data-stu-id="becf2-117">Minimum</span></span><br><span data-ttu-id="becf2-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="becf2-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="becf2-119">context</span><span class="sxs-lookup"><span data-stu-id="becf2-119">context</span></span>](office.context.md) | <span data-ttu-id="becf2-120">作成</span><span class="sxs-lookup"><span data-stu-id="becf2-120">Compose</span></span><br><span data-ttu-id="becf2-121">読み取り</span><span class="sxs-lookup"><span data-stu-id="becf2-121">Read</span></span> | [<span data-ttu-id="becf2-122">Context</span><span class="sxs-lookup"><span data-stu-id="becf2-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview) | [<span data-ttu-id="becf2-123">1.1</span><span class="sxs-lookup"><span data-stu-id="becf2-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="becf2-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="becf2-124">Enumerations</span></span>

| <span data-ttu-id="becf2-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="becf2-125">Enumeration</span></span> | <span data-ttu-id="becf2-126">モード</span><span class="sxs-lookup"><span data-stu-id="becf2-126">Modes</span></span> | <span data-ttu-id="becf2-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="becf2-127">Return type</span></span> | <span data-ttu-id="becf2-128">最小値</span><span class="sxs-lookup"><span data-stu-id="becf2-128">Minimum</span></span><br><span data-ttu-id="becf2-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="becf2-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="becf2-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="becf2-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="becf2-131">作成</span><span class="sxs-lookup"><span data-stu-id="becf2-131">Compose</span></span><br><span data-ttu-id="becf2-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="becf2-132">Read</span></span> | <span data-ttu-id="becf2-133">文字列</span><span class="sxs-lookup"><span data-stu-id="becf2-133">String</span></span> | [<span data-ttu-id="becf2-134">1.1</span><span class="sxs-lookup"><span data-stu-id="becf2-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="becf2-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="becf2-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="becf2-136">作成</span><span class="sxs-lookup"><span data-stu-id="becf2-136">Compose</span></span><br><span data-ttu-id="becf2-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="becf2-137">Read</span></span> | <span data-ttu-id="becf2-138">文字列</span><span class="sxs-lookup"><span data-stu-id="becf2-138">String</span></span> | [<span data-ttu-id="becf2-139">1.1</span><span class="sxs-lookup"><span data-stu-id="becf2-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="becf2-140">EventType</span><span class="sxs-lookup"><span data-stu-id="becf2-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="becf2-141">作成</span><span class="sxs-lookup"><span data-stu-id="becf2-141">Compose</span></span><br><span data-ttu-id="becf2-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="becf2-142">Read</span></span> | <span data-ttu-id="becf2-143">文字列</span><span class="sxs-lookup"><span data-stu-id="becf2-143">String</span></span> | [<span data-ttu-id="becf2-144">1.5</span><span class="sxs-lookup"><span data-stu-id="becf2-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="becf2-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="becf2-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="becf2-146">作成</span><span class="sxs-lookup"><span data-stu-id="becf2-146">Compose</span></span><br><span data-ttu-id="becf2-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="becf2-147">Read</span></span> | <span data-ttu-id="becf2-148">文字列</span><span class="sxs-lookup"><span data-stu-id="becf2-148">String</span></span> | [<span data-ttu-id="becf2-149">1.1</span><span class="sxs-lookup"><span data-stu-id="becf2-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="becf2-150">名前空間</span><span class="sxs-lookup"><span data-stu-id="becf2-150">Namespaces</span></span>

<span data-ttu-id="becf2-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="becf2-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="becf2-152">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="becf2-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="becf2-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="becf2-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="becf2-154">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="becf2-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="becf2-155">型</span><span class="sxs-lookup"><span data-stu-id="becf2-155">Type</span></span>

*   <span data-ttu-id="becf2-156">String</span><span class="sxs-lookup"><span data-stu-id="becf2-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="becf2-157">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="becf2-157">Properties:</span></span>

|<span data-ttu-id="becf2-158">名前</span><span class="sxs-lookup"><span data-stu-id="becf2-158">Name</span></span>| <span data-ttu-id="becf2-159">種類</span><span class="sxs-lookup"><span data-stu-id="becf2-159">Type</span></span>| <span data-ttu-id="becf2-160">説明</span><span class="sxs-lookup"><span data-stu-id="becf2-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="becf2-161">文字列</span><span class="sxs-lookup"><span data-stu-id="becf2-161">String</span></span>|<span data-ttu-id="becf2-162">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="becf2-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="becf2-163">文字列</span><span class="sxs-lookup"><span data-stu-id="becf2-163">String</span></span>|<span data-ttu-id="becf2-164">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="becf2-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="becf2-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="becf2-165">Requirements</span></span>

|<span data-ttu-id="becf2-166">要件</span><span class="sxs-lookup"><span data-stu-id="becf2-166">Requirement</span></span>| <span data-ttu-id="becf2-167">値</span><span class="sxs-lookup"><span data-stu-id="becf2-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="becf2-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="becf2-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="becf2-169">1.1</span><span class="sxs-lookup"><span data-stu-id="becf2-169">1.1</span></span>|
|[<span data-ttu-id="becf2-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="becf2-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="becf2-171">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="becf2-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="becf2-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="becf2-172">CoercionType: String</span></span>

<span data-ttu-id="becf2-173">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="becf2-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="becf2-174">型</span><span class="sxs-lookup"><span data-stu-id="becf2-174">Type</span></span>

*   <span data-ttu-id="becf2-175">String</span><span class="sxs-lookup"><span data-stu-id="becf2-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="becf2-176">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="becf2-176">Properties:</span></span>

|<span data-ttu-id="becf2-177">名前</span><span class="sxs-lookup"><span data-stu-id="becf2-177">Name</span></span>| <span data-ttu-id="becf2-178">種類</span><span class="sxs-lookup"><span data-stu-id="becf2-178">Type</span></span>| <span data-ttu-id="becf2-179">説明</span><span class="sxs-lookup"><span data-stu-id="becf2-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="becf2-180">文字列</span><span class="sxs-lookup"><span data-stu-id="becf2-180">String</span></span>|<span data-ttu-id="becf2-181">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="becf2-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="becf2-182">String</span><span class="sxs-lookup"><span data-stu-id="becf2-182">String</span></span>|<span data-ttu-id="becf2-183">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="becf2-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="becf2-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="becf2-184">Requirements</span></span>

|<span data-ttu-id="becf2-185">要件</span><span class="sxs-lookup"><span data-stu-id="becf2-185">Requirement</span></span>| <span data-ttu-id="becf2-186">値</span><span class="sxs-lookup"><span data-stu-id="becf2-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="becf2-187">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="becf2-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="becf2-188">1.1</span><span class="sxs-lookup"><span data-stu-id="becf2-188">1.1</span></span>|
|[<span data-ttu-id="becf2-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="becf2-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="becf2-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="becf2-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="becf2-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="becf2-191">EventType: String</span></span>

<span data-ttu-id="becf2-192">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="becf2-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="becf2-193">型</span><span class="sxs-lookup"><span data-stu-id="becf2-193">Type</span></span>

*   <span data-ttu-id="becf2-194">String</span><span class="sxs-lookup"><span data-stu-id="becf2-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="becf2-195">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="becf2-195">Properties:</span></span>

| <span data-ttu-id="becf2-196">名前</span><span class="sxs-lookup"><span data-stu-id="becf2-196">Name</span></span> | <span data-ttu-id="becf2-197">種類</span><span class="sxs-lookup"><span data-stu-id="becf2-197">Type</span></span> | <span data-ttu-id="becf2-198">説明</span><span class="sxs-lookup"><span data-stu-id="becf2-198">Description</span></span> | <span data-ttu-id="becf2-199">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="becf2-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="becf2-200">文字列</span><span class="sxs-lookup"><span data-stu-id="becf2-200">String</span></span> | <span data-ttu-id="becf2-201">選択した予定またはデータ系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="becf2-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="becf2-202">1.7</span><span class="sxs-lookup"><span data-stu-id="becf2-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="becf2-203">文字列</span><span class="sxs-lookup"><span data-stu-id="becf2-203">String</span></span> | <span data-ttu-id="becf2-204">添付ファイルがアイテムに追加またはアイテムから削除されています。</span><span class="sxs-lookup"><span data-stu-id="becf2-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="becf2-205">1.8</span><span class="sxs-lookup"><span data-stu-id="becf2-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="becf2-206">文字列</span><span class="sxs-lookup"><span data-stu-id="becf2-206">String</span></span> | <span data-ttu-id="becf2-207">選択した予定の場所が変更されました。</span><span class="sxs-lookup"><span data-stu-id="becf2-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="becf2-208">1.8</span><span class="sxs-lookup"><span data-stu-id="becf2-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="becf2-209">文字列</span><span class="sxs-lookup"><span data-stu-id="becf2-209">String</span></span> | <span data-ttu-id="becf2-210">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="becf2-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="becf2-211">1.5</span><span class="sxs-lookup"><span data-stu-id="becf2-211">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="becf2-212">文字列</span><span class="sxs-lookup"><span data-stu-id="becf2-212">String</span></span> | <span data-ttu-id="becf2-213">メールボックスの Office テーマが変更されました。</span><span class="sxs-lookup"><span data-stu-id="becf2-213">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="becf2-214">プレビュー</span><span class="sxs-lookup"><span data-stu-id="becf2-214">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="becf2-215">文字列</span><span class="sxs-lookup"><span data-stu-id="becf2-215">String</span></span> | <span data-ttu-id="becf2-216">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="becf2-216">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="becf2-217">1.7</span><span class="sxs-lookup"><span data-stu-id="becf2-217">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="becf2-218">文字列</span><span class="sxs-lookup"><span data-stu-id="becf2-218">String</span></span> | <span data-ttu-id="becf2-219">選択したアイテムの定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="becf2-219">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="becf2-220">1.7</span><span class="sxs-lookup"><span data-stu-id="becf2-220">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="becf2-221">Requirements</span><span class="sxs-lookup"><span data-stu-id="becf2-221">Requirements</span></span>

|<span data-ttu-id="becf2-222">要件</span><span class="sxs-lookup"><span data-stu-id="becf2-222">Requirement</span></span>| <span data-ttu-id="becf2-223">値</span><span class="sxs-lookup"><span data-stu-id="becf2-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="becf2-224">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="becf2-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="becf2-225">1.5</span><span class="sxs-lookup"><span data-stu-id="becf2-225">1.5</span></span> |
|[<span data-ttu-id="becf2-226">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="becf2-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="becf2-227">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="becf2-227">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="becf2-228">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="becf2-228">SourceProperty: String</span></span>

<span data-ttu-id="becf2-229">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="becf2-229">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="becf2-230">型</span><span class="sxs-lookup"><span data-stu-id="becf2-230">Type</span></span>

*   <span data-ttu-id="becf2-231">String</span><span class="sxs-lookup"><span data-stu-id="becf2-231">String</span></span>

##### <a name="properties"></a><span data-ttu-id="becf2-232">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="becf2-232">Properties:</span></span>

|<span data-ttu-id="becf2-233">名前</span><span class="sxs-lookup"><span data-stu-id="becf2-233">Name</span></span>| <span data-ttu-id="becf2-234">種類</span><span class="sxs-lookup"><span data-stu-id="becf2-234">Type</span></span>| <span data-ttu-id="becf2-235">説明</span><span class="sxs-lookup"><span data-stu-id="becf2-235">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="becf2-236">文字列</span><span class="sxs-lookup"><span data-stu-id="becf2-236">String</span></span>|<span data-ttu-id="becf2-237">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="becf2-237">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="becf2-238">文字列</span><span class="sxs-lookup"><span data-stu-id="becf2-238">String</span></span>|<span data-ttu-id="becf2-239">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="becf2-239">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="becf2-240">Requirements</span><span class="sxs-lookup"><span data-stu-id="becf2-240">Requirements</span></span>

|<span data-ttu-id="becf2-241">要件</span><span class="sxs-lookup"><span data-stu-id="becf2-241">Requirement</span></span>| <span data-ttu-id="becf2-242">値</span><span class="sxs-lookup"><span data-stu-id="becf2-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="becf2-243">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="becf2-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="becf2-244">1.1</span><span class="sxs-lookup"><span data-stu-id="becf2-244">1.1</span></span>|
|[<span data-ttu-id="becf2-245">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="becf2-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="becf2-246">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="becf2-246">Compose or Read</span></span>|
