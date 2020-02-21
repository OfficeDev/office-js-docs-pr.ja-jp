---
title: Office 名前空間-要件セット1.7
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 23f3fb705c03eabd8ee7fce53f4c89a48128672f
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165349"
---
# <a name="office"></a><span data-ttu-id="04aa1-102">Office</span><span class="sxs-lookup"><span data-stu-id="04aa1-102">Office</span></span>

<span data-ttu-id="04aa1-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="04aa1-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="04aa1-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="04aa1-105">Requirements</span></span>

|<span data-ttu-id="04aa1-106">要件</span><span class="sxs-lookup"><span data-stu-id="04aa1-106">Requirement</span></span>| <span data-ttu-id="04aa1-107">値</span><span class="sxs-lookup"><span data-stu-id="04aa1-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="04aa1-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="04aa1-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="04aa1-109">1.1</span><span class="sxs-lookup"><span data-stu-id="04aa1-109">1.1</span></span>|
|[<span data-ttu-id="04aa1-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="04aa1-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="04aa1-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="04aa1-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="04aa1-112">Properties</span><span class="sxs-lookup"><span data-stu-id="04aa1-112">Properties</span></span>

| <span data-ttu-id="04aa1-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="04aa1-113">Property</span></span> | <span data-ttu-id="04aa1-114">モード</span><span class="sxs-lookup"><span data-stu-id="04aa1-114">Modes</span></span> | <span data-ttu-id="04aa1-115">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="04aa1-115">Return type</span></span> | <span data-ttu-id="04aa1-116">最小値</span><span class="sxs-lookup"><span data-stu-id="04aa1-116">Minimum</span></span><br><span data-ttu-id="04aa1-117">要件セット</span><span class="sxs-lookup"><span data-stu-id="04aa1-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="04aa1-118">context</span><span class="sxs-lookup"><span data-stu-id="04aa1-118">context</span></span>](office.context.md) | <span data-ttu-id="04aa1-119">作成</span><span class="sxs-lookup"><span data-stu-id="04aa1-119">Compose</span></span><br><span data-ttu-id="04aa1-120">読み取り</span><span class="sxs-lookup"><span data-stu-id="04aa1-120">Read</span></span> | [<span data-ttu-id="04aa1-121">Context</span><span class="sxs-lookup"><span data-stu-id="04aa1-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7) | [<span data-ttu-id="04aa1-122">1.1</span><span class="sxs-lookup"><span data-stu-id="04aa1-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="04aa1-123">列挙型</span><span class="sxs-lookup"><span data-stu-id="04aa1-123">Enumerations</span></span>

| <span data-ttu-id="04aa1-124">列挙体</span><span class="sxs-lookup"><span data-stu-id="04aa1-124">Enumeration</span></span> | <span data-ttu-id="04aa1-125">モード</span><span class="sxs-lookup"><span data-stu-id="04aa1-125">Modes</span></span> | <span data-ttu-id="04aa1-126">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="04aa1-126">Return type</span></span> | <span data-ttu-id="04aa1-127">最小値</span><span class="sxs-lookup"><span data-stu-id="04aa1-127">Minimum</span></span><br><span data-ttu-id="04aa1-128">要件セット</span><span class="sxs-lookup"><span data-stu-id="04aa1-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="04aa1-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="04aa1-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="04aa1-130">作成</span><span class="sxs-lookup"><span data-stu-id="04aa1-130">Compose</span></span><br><span data-ttu-id="04aa1-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="04aa1-131">Read</span></span> | <span data-ttu-id="04aa1-132">文字列</span><span class="sxs-lookup"><span data-stu-id="04aa1-132">String</span></span> | [<span data-ttu-id="04aa1-133">1.1</span><span class="sxs-lookup"><span data-stu-id="04aa1-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="04aa1-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="04aa1-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="04aa1-135">作成</span><span class="sxs-lookup"><span data-stu-id="04aa1-135">Compose</span></span><br><span data-ttu-id="04aa1-136">読み取り</span><span class="sxs-lookup"><span data-stu-id="04aa1-136">Read</span></span> | <span data-ttu-id="04aa1-137">文字列</span><span class="sxs-lookup"><span data-stu-id="04aa1-137">String</span></span> | [<span data-ttu-id="04aa1-138">1.1</span><span class="sxs-lookup"><span data-stu-id="04aa1-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="04aa1-139">EventType</span><span class="sxs-lookup"><span data-stu-id="04aa1-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="04aa1-140">作成</span><span class="sxs-lookup"><span data-stu-id="04aa1-140">Compose</span></span><br><span data-ttu-id="04aa1-141">読み取り</span><span class="sxs-lookup"><span data-stu-id="04aa1-141">Read</span></span> | <span data-ttu-id="04aa1-142">文字列</span><span class="sxs-lookup"><span data-stu-id="04aa1-142">String</span></span> | [<span data-ttu-id="04aa1-143">1.5</span><span class="sxs-lookup"><span data-stu-id="04aa1-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="04aa1-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="04aa1-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="04aa1-145">作成</span><span class="sxs-lookup"><span data-stu-id="04aa1-145">Compose</span></span><br><span data-ttu-id="04aa1-146">読み取り</span><span class="sxs-lookup"><span data-stu-id="04aa1-146">Read</span></span> | <span data-ttu-id="04aa1-147">文字列</span><span class="sxs-lookup"><span data-stu-id="04aa1-147">String</span></span> | [<span data-ttu-id="04aa1-148">1.1</span><span class="sxs-lookup"><span data-stu-id="04aa1-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="04aa1-149">名前空間</span><span class="sxs-lookup"><span data-stu-id="04aa1-149">Namespaces</span></span>

<span data-ttu-id="04aa1-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="04aa1-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="04aa1-151">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="04aa1-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="04aa1-152">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="04aa1-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="04aa1-153">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="04aa1-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="04aa1-154">型</span><span class="sxs-lookup"><span data-stu-id="04aa1-154">Type</span></span>

*   <span data-ttu-id="04aa1-155">String</span><span class="sxs-lookup"><span data-stu-id="04aa1-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="04aa1-156">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="04aa1-156">Properties:</span></span>

|<span data-ttu-id="04aa1-157">名前</span><span class="sxs-lookup"><span data-stu-id="04aa1-157">Name</span></span>| <span data-ttu-id="04aa1-158">種類</span><span class="sxs-lookup"><span data-stu-id="04aa1-158">Type</span></span>| <span data-ttu-id="04aa1-159">説明</span><span class="sxs-lookup"><span data-stu-id="04aa1-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="04aa1-160">文字列</span><span class="sxs-lookup"><span data-stu-id="04aa1-160">String</span></span>|<span data-ttu-id="04aa1-161">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="04aa1-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="04aa1-162">文字列</span><span class="sxs-lookup"><span data-stu-id="04aa1-162">String</span></span>|<span data-ttu-id="04aa1-163">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="04aa1-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="04aa1-164">Requirements</span><span class="sxs-lookup"><span data-stu-id="04aa1-164">Requirements</span></span>

|<span data-ttu-id="04aa1-165">要件</span><span class="sxs-lookup"><span data-stu-id="04aa1-165">Requirement</span></span>| <span data-ttu-id="04aa1-166">値</span><span class="sxs-lookup"><span data-stu-id="04aa1-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="04aa1-167">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="04aa1-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="04aa1-168">1.1</span><span class="sxs-lookup"><span data-stu-id="04aa1-168">1.1</span></span>|
|[<span data-ttu-id="04aa1-169">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="04aa1-169">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="04aa1-170">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="04aa1-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="04aa1-171">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="04aa1-171">CoercionType: String</span></span>

<span data-ttu-id="04aa1-172">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="04aa1-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="04aa1-173">型</span><span class="sxs-lookup"><span data-stu-id="04aa1-173">Type</span></span>

*   <span data-ttu-id="04aa1-174">String</span><span class="sxs-lookup"><span data-stu-id="04aa1-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="04aa1-175">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="04aa1-175">Properties:</span></span>

|<span data-ttu-id="04aa1-176">名前</span><span class="sxs-lookup"><span data-stu-id="04aa1-176">Name</span></span>| <span data-ttu-id="04aa1-177">種類</span><span class="sxs-lookup"><span data-stu-id="04aa1-177">Type</span></span>| <span data-ttu-id="04aa1-178">説明</span><span class="sxs-lookup"><span data-stu-id="04aa1-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="04aa1-179">文字列</span><span class="sxs-lookup"><span data-stu-id="04aa1-179">String</span></span>|<span data-ttu-id="04aa1-180">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="04aa1-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="04aa1-181">String</span><span class="sxs-lookup"><span data-stu-id="04aa1-181">String</span></span>|<span data-ttu-id="04aa1-182">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="04aa1-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="04aa1-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="04aa1-183">Requirements</span></span>

|<span data-ttu-id="04aa1-184">要件</span><span class="sxs-lookup"><span data-stu-id="04aa1-184">Requirement</span></span>| <span data-ttu-id="04aa1-185">値</span><span class="sxs-lookup"><span data-stu-id="04aa1-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="04aa1-186">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="04aa1-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="04aa1-187">1.1</span><span class="sxs-lookup"><span data-stu-id="04aa1-187">1.1</span></span>|
|[<span data-ttu-id="04aa1-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="04aa1-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="04aa1-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="04aa1-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="04aa1-190">EventType: String</span><span class="sxs-lookup"><span data-stu-id="04aa1-190">EventType: String</span></span>

<span data-ttu-id="04aa1-191">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="04aa1-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="04aa1-192">型</span><span class="sxs-lookup"><span data-stu-id="04aa1-192">Type</span></span>

*   <span data-ttu-id="04aa1-193">String</span><span class="sxs-lookup"><span data-stu-id="04aa1-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="04aa1-194">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="04aa1-194">Properties:</span></span>

| <span data-ttu-id="04aa1-195">名前</span><span class="sxs-lookup"><span data-stu-id="04aa1-195">Name</span></span> | <span data-ttu-id="04aa1-196">種類</span><span class="sxs-lookup"><span data-stu-id="04aa1-196">Type</span></span> | <span data-ttu-id="04aa1-197">説明</span><span class="sxs-lookup"><span data-stu-id="04aa1-197">Description</span></span> | <span data-ttu-id="04aa1-198">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="04aa1-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="04aa1-199">文字列</span><span class="sxs-lookup"><span data-stu-id="04aa1-199">String</span></span> | <span data-ttu-id="04aa1-200">選択した予定またはデータ系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="04aa1-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="04aa1-201">1.7</span><span class="sxs-lookup"><span data-stu-id="04aa1-201">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="04aa1-202">文字列</span><span class="sxs-lookup"><span data-stu-id="04aa1-202">String</span></span> | <span data-ttu-id="04aa1-203">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="04aa1-203">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="04aa1-204">1.5</span><span class="sxs-lookup"><span data-stu-id="04aa1-204">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="04aa1-205">文字列</span><span class="sxs-lookup"><span data-stu-id="04aa1-205">String</span></span> | <span data-ttu-id="04aa1-206">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="04aa1-206">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="04aa1-207">1.7</span><span class="sxs-lookup"><span data-stu-id="04aa1-207">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="04aa1-208">文字列</span><span class="sxs-lookup"><span data-stu-id="04aa1-208">String</span></span> | <span data-ttu-id="04aa1-209">選択したアイテムの定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="04aa1-209">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="04aa1-210">1.7</span><span class="sxs-lookup"><span data-stu-id="04aa1-210">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="04aa1-211">Requirements</span><span class="sxs-lookup"><span data-stu-id="04aa1-211">Requirements</span></span>

|<span data-ttu-id="04aa1-212">要件</span><span class="sxs-lookup"><span data-stu-id="04aa1-212">Requirement</span></span>| <span data-ttu-id="04aa1-213">値</span><span class="sxs-lookup"><span data-stu-id="04aa1-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="04aa1-214">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="04aa1-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="04aa1-215">1.5</span><span class="sxs-lookup"><span data-stu-id="04aa1-215">1.5</span></span> |
|[<span data-ttu-id="04aa1-216">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="04aa1-216">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="04aa1-217">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="04aa1-217">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="04aa1-218">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="04aa1-218">SourceProperty: String</span></span>

<span data-ttu-id="04aa1-219">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="04aa1-219">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="04aa1-220">型</span><span class="sxs-lookup"><span data-stu-id="04aa1-220">Type</span></span>

*   <span data-ttu-id="04aa1-221">String</span><span class="sxs-lookup"><span data-stu-id="04aa1-221">String</span></span>

##### <a name="properties"></a><span data-ttu-id="04aa1-222">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="04aa1-222">Properties:</span></span>

|<span data-ttu-id="04aa1-223">名前</span><span class="sxs-lookup"><span data-stu-id="04aa1-223">Name</span></span>| <span data-ttu-id="04aa1-224">種類</span><span class="sxs-lookup"><span data-stu-id="04aa1-224">Type</span></span>| <span data-ttu-id="04aa1-225">説明</span><span class="sxs-lookup"><span data-stu-id="04aa1-225">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="04aa1-226">文字列</span><span class="sxs-lookup"><span data-stu-id="04aa1-226">String</span></span>|<span data-ttu-id="04aa1-227">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="04aa1-227">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="04aa1-228">文字列</span><span class="sxs-lookup"><span data-stu-id="04aa1-228">String</span></span>|<span data-ttu-id="04aa1-229">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="04aa1-229">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="04aa1-230">Requirements</span><span class="sxs-lookup"><span data-stu-id="04aa1-230">Requirements</span></span>

|<span data-ttu-id="04aa1-231">要件</span><span class="sxs-lookup"><span data-stu-id="04aa1-231">Requirement</span></span>| <span data-ttu-id="04aa1-232">値</span><span class="sxs-lookup"><span data-stu-id="04aa1-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="04aa1-233">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="04aa1-233">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="04aa1-234">1.1</span><span class="sxs-lookup"><span data-stu-id="04aa1-234">1.1</span></span>|
|[<span data-ttu-id="04aa1-235">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="04aa1-235">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="04aa1-236">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="04aa1-236">Compose or Read</span></span>|
