---
title: Office 名前空間-要件セット1.7
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 9bfff9c45cb157d2dcd42997a01f5ada40aecfa0
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814571"
---
# <a name="office"></a><span data-ttu-id="6f89e-102">Office</span><span class="sxs-lookup"><span data-stu-id="6f89e-102">Office</span></span>

<span data-ttu-id="6f89e-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6f89e-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="6f89e-105">要件</span><span class="sxs-lookup"><span data-stu-id="6f89e-105">Requirements</span></span>

|<span data-ttu-id="6f89e-106">要件</span><span class="sxs-lookup"><span data-stu-id="6f89e-106">Requirement</span></span>| <span data-ttu-id="6f89e-107">値</span><span class="sxs-lookup"><span data-stu-id="6f89e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="6f89e-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6f89e-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6f89e-109">1.1</span><span class="sxs-lookup"><span data-stu-id="6f89e-109">1.1</span></span>|
|[<span data-ttu-id="6f89e-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6f89e-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6f89e-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6f89e-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="6f89e-112">Properties</span><span class="sxs-lookup"><span data-stu-id="6f89e-112">Properties</span></span>

| <span data-ttu-id="6f89e-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="6f89e-113">Property</span></span> | <span data-ttu-id="6f89e-114">モード</span><span class="sxs-lookup"><span data-stu-id="6f89e-114">Modes</span></span> | <span data-ttu-id="6f89e-115">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="6f89e-115">Return type</span></span> | <span data-ttu-id="6f89e-116">最小値</span><span class="sxs-lookup"><span data-stu-id="6f89e-116">Minimum</span></span><br><span data-ttu-id="6f89e-117">要件セット</span><span class="sxs-lookup"><span data-stu-id="6f89e-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="6f89e-118">context</span><span class="sxs-lookup"><span data-stu-id="6f89e-118">context</span></span>](office.context.md) | <span data-ttu-id="6f89e-119">作成</span><span class="sxs-lookup"><span data-stu-id="6f89e-119">Compose</span></span><br><span data-ttu-id="6f89e-120">読み取り</span><span class="sxs-lookup"><span data-stu-id="6f89e-120">Read</span></span> | [<span data-ttu-id="6f89e-121">Context</span><span class="sxs-lookup"><span data-stu-id="6f89e-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7) | [<span data-ttu-id="6f89e-122">1.1</span><span class="sxs-lookup"><span data-stu-id="6f89e-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="6f89e-123">列挙型</span><span class="sxs-lookup"><span data-stu-id="6f89e-123">Enumerations</span></span>

| <span data-ttu-id="6f89e-124">列挙体</span><span class="sxs-lookup"><span data-stu-id="6f89e-124">Enumeration</span></span> | <span data-ttu-id="6f89e-125">モード</span><span class="sxs-lookup"><span data-stu-id="6f89e-125">Modes</span></span> | <span data-ttu-id="6f89e-126">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="6f89e-126">Return type</span></span> | <span data-ttu-id="6f89e-127">最小値</span><span class="sxs-lookup"><span data-stu-id="6f89e-127">Minimum</span></span><br><span data-ttu-id="6f89e-128">要件セット</span><span class="sxs-lookup"><span data-stu-id="6f89e-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="6f89e-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="6f89e-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="6f89e-130">作成</span><span class="sxs-lookup"><span data-stu-id="6f89e-130">Compose</span></span><br><span data-ttu-id="6f89e-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="6f89e-131">Read</span></span> | <span data-ttu-id="6f89e-132">String</span><span class="sxs-lookup"><span data-stu-id="6f89e-132">String</span></span> | [<span data-ttu-id="6f89e-133">1.1</span><span class="sxs-lookup"><span data-stu-id="6f89e-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6f89e-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="6f89e-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="6f89e-135">作成</span><span class="sxs-lookup"><span data-stu-id="6f89e-135">Compose</span></span><br><span data-ttu-id="6f89e-136">読み取り</span><span class="sxs-lookup"><span data-stu-id="6f89e-136">Read</span></span> | <span data-ttu-id="6f89e-137">String</span><span class="sxs-lookup"><span data-stu-id="6f89e-137">String</span></span> | [<span data-ttu-id="6f89e-138">1.1</span><span class="sxs-lookup"><span data-stu-id="6f89e-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6f89e-139">EventType</span><span class="sxs-lookup"><span data-stu-id="6f89e-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="6f89e-140">作成</span><span class="sxs-lookup"><span data-stu-id="6f89e-140">Compose</span></span><br><span data-ttu-id="6f89e-141">読み取り</span><span class="sxs-lookup"><span data-stu-id="6f89e-141">Read</span></span> | <span data-ttu-id="6f89e-142">String</span><span class="sxs-lookup"><span data-stu-id="6f89e-142">String</span></span> | [<span data-ttu-id="6f89e-143">1.5</span><span class="sxs-lookup"><span data-stu-id="6f89e-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="6f89e-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="6f89e-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="6f89e-145">作成</span><span class="sxs-lookup"><span data-stu-id="6f89e-145">Compose</span></span><br><span data-ttu-id="6f89e-146">読み取り</span><span class="sxs-lookup"><span data-stu-id="6f89e-146">Read</span></span> | <span data-ttu-id="6f89e-147">String</span><span class="sxs-lookup"><span data-stu-id="6f89e-147">String</span></span> | [<span data-ttu-id="6f89e-148">1.1</span><span class="sxs-lookup"><span data-stu-id="6f89e-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="6f89e-149">名前空間</span><span class="sxs-lookup"><span data-stu-id="6f89e-149">Namespaces</span></span>

<span data-ttu-id="6f89e-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="6f89e-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="6f89e-151">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="6f89e-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="6f89e-152">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="6f89e-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="6f89e-153">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="6f89e-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="6f89e-154">型</span><span class="sxs-lookup"><span data-stu-id="6f89e-154">Type</span></span>

*   <span data-ttu-id="6f89e-155">String</span><span class="sxs-lookup"><span data-stu-id="6f89e-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6f89e-156">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="6f89e-156">Properties:</span></span>

|<span data-ttu-id="6f89e-157">名前</span><span class="sxs-lookup"><span data-stu-id="6f89e-157">Name</span></span>| <span data-ttu-id="6f89e-158">種類</span><span class="sxs-lookup"><span data-stu-id="6f89e-158">Type</span></span>| <span data-ttu-id="6f89e-159">説明</span><span class="sxs-lookup"><span data-stu-id="6f89e-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="6f89e-160">String</span><span class="sxs-lookup"><span data-stu-id="6f89e-160">String</span></span>|<span data-ttu-id="6f89e-161">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="6f89e-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="6f89e-162">String</span><span class="sxs-lookup"><span data-stu-id="6f89e-162">String</span></span>|<span data-ttu-id="6f89e-163">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="6f89e-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6f89e-164">要件</span><span class="sxs-lookup"><span data-stu-id="6f89e-164">Requirements</span></span>

|<span data-ttu-id="6f89e-165">要件</span><span class="sxs-lookup"><span data-stu-id="6f89e-165">Requirement</span></span>| <span data-ttu-id="6f89e-166">値</span><span class="sxs-lookup"><span data-stu-id="6f89e-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="6f89e-167">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6f89e-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6f89e-168">1.1</span><span class="sxs-lookup"><span data-stu-id="6f89e-168">1.1</span></span>|
|[<span data-ttu-id="6f89e-169">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6f89e-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6f89e-170">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6f89e-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="6f89e-171">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="6f89e-171">CoercionType: String</span></span>

<span data-ttu-id="6f89e-172">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="6f89e-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6f89e-173">型</span><span class="sxs-lookup"><span data-stu-id="6f89e-173">Type</span></span>

*   <span data-ttu-id="6f89e-174">String</span><span class="sxs-lookup"><span data-stu-id="6f89e-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6f89e-175">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="6f89e-175">Properties:</span></span>

|<span data-ttu-id="6f89e-176">名前</span><span class="sxs-lookup"><span data-stu-id="6f89e-176">Name</span></span>| <span data-ttu-id="6f89e-177">種類</span><span class="sxs-lookup"><span data-stu-id="6f89e-177">Type</span></span>| <span data-ttu-id="6f89e-178">説明</span><span class="sxs-lookup"><span data-stu-id="6f89e-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="6f89e-179">String</span><span class="sxs-lookup"><span data-stu-id="6f89e-179">String</span></span>|<span data-ttu-id="6f89e-180">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="6f89e-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="6f89e-181">String</span><span class="sxs-lookup"><span data-stu-id="6f89e-181">String</span></span>|<span data-ttu-id="6f89e-182">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="6f89e-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6f89e-183">要件</span><span class="sxs-lookup"><span data-stu-id="6f89e-183">Requirements</span></span>

|<span data-ttu-id="6f89e-184">要件</span><span class="sxs-lookup"><span data-stu-id="6f89e-184">Requirement</span></span>| <span data-ttu-id="6f89e-185">値</span><span class="sxs-lookup"><span data-stu-id="6f89e-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="6f89e-186">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6f89e-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6f89e-187">1.1</span><span class="sxs-lookup"><span data-stu-id="6f89e-187">1.1</span></span>|
|[<span data-ttu-id="6f89e-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6f89e-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6f89e-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6f89e-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="6f89e-190">EventType: String</span><span class="sxs-lookup"><span data-stu-id="6f89e-190">EventType: String</span></span>

<span data-ttu-id="6f89e-191">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="6f89e-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="6f89e-192">型</span><span class="sxs-lookup"><span data-stu-id="6f89e-192">Type</span></span>

*   <span data-ttu-id="6f89e-193">String</span><span class="sxs-lookup"><span data-stu-id="6f89e-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6f89e-194">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="6f89e-194">Properties:</span></span>

| <span data-ttu-id="6f89e-195">名前</span><span class="sxs-lookup"><span data-stu-id="6f89e-195">Name</span></span> | <span data-ttu-id="6f89e-196">種類</span><span class="sxs-lookup"><span data-stu-id="6f89e-196">Type</span></span> | <span data-ttu-id="6f89e-197">説明</span><span class="sxs-lookup"><span data-stu-id="6f89e-197">Description</span></span> | <span data-ttu-id="6f89e-198">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="6f89e-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="6f89e-199">String</span><span class="sxs-lookup"><span data-stu-id="6f89e-199">String</span></span> | <span data-ttu-id="6f89e-200">選択した予定またはデータ系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="6f89e-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="6f89e-201">1.7</span><span class="sxs-lookup"><span data-stu-id="6f89e-201">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="6f89e-202">String</span><span class="sxs-lookup"><span data-stu-id="6f89e-202">String</span></span> | <span data-ttu-id="6f89e-203">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="6f89e-203">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="6f89e-204">1.5</span><span class="sxs-lookup"><span data-stu-id="6f89e-204">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="6f89e-205">String</span><span class="sxs-lookup"><span data-stu-id="6f89e-205">String</span></span> | <span data-ttu-id="6f89e-206">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="6f89e-206">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="6f89e-207">1.7</span><span class="sxs-lookup"><span data-stu-id="6f89e-207">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="6f89e-208">String</span><span class="sxs-lookup"><span data-stu-id="6f89e-208">String</span></span> | <span data-ttu-id="6f89e-209">選択したアイテムの定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="6f89e-209">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="6f89e-210">1.7</span><span class="sxs-lookup"><span data-stu-id="6f89e-210">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6f89e-211">要件</span><span class="sxs-lookup"><span data-stu-id="6f89e-211">Requirements</span></span>

|<span data-ttu-id="6f89e-212">要件</span><span class="sxs-lookup"><span data-stu-id="6f89e-212">Requirement</span></span>| <span data-ttu-id="6f89e-213">値</span><span class="sxs-lookup"><span data-stu-id="6f89e-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="6f89e-214">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6f89e-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6f89e-215">1.5</span><span class="sxs-lookup"><span data-stu-id="6f89e-215">1.5</span></span> |
|[<span data-ttu-id="6f89e-216">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6f89e-216">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6f89e-217">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6f89e-217">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="6f89e-218">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="6f89e-218">SourceProperty: String</span></span>

<span data-ttu-id="6f89e-219">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="6f89e-219">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6f89e-220">型</span><span class="sxs-lookup"><span data-stu-id="6f89e-220">Type</span></span>

*   <span data-ttu-id="6f89e-221">String</span><span class="sxs-lookup"><span data-stu-id="6f89e-221">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6f89e-222">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="6f89e-222">Properties:</span></span>

|<span data-ttu-id="6f89e-223">名前</span><span class="sxs-lookup"><span data-stu-id="6f89e-223">Name</span></span>| <span data-ttu-id="6f89e-224">種類</span><span class="sxs-lookup"><span data-stu-id="6f89e-224">Type</span></span>| <span data-ttu-id="6f89e-225">説明</span><span class="sxs-lookup"><span data-stu-id="6f89e-225">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="6f89e-226">String</span><span class="sxs-lookup"><span data-stu-id="6f89e-226">String</span></span>|<span data-ttu-id="6f89e-227">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="6f89e-227">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="6f89e-228">String</span><span class="sxs-lookup"><span data-stu-id="6f89e-228">String</span></span>|<span data-ttu-id="6f89e-229">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="6f89e-229">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6f89e-230">要件</span><span class="sxs-lookup"><span data-stu-id="6f89e-230">Requirements</span></span>

|<span data-ttu-id="6f89e-231">要件</span><span class="sxs-lookup"><span data-stu-id="6f89e-231">Requirement</span></span>| <span data-ttu-id="6f89e-232">値</span><span class="sxs-lookup"><span data-stu-id="6f89e-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="6f89e-233">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6f89e-233">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6f89e-234">1.1</span><span class="sxs-lookup"><span data-stu-id="6f89e-234">1.1</span></span>|
|[<span data-ttu-id="6f89e-235">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6f89e-235">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6f89e-236">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6f89e-236">Compose or Read</span></span>|
