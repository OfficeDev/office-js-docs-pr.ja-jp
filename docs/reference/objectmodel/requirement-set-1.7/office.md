---
title: Office 名前空間-要件セット1.7
description: この名前空間は、Outlook Office アドインで使用される共有インターフェイスを提供します (要件セット 1.7)
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 50fa22ac14aee3b7276be83813db248681435dc1
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717599"
---
# <a name="office"></a><span data-ttu-id="2d3b4-103">Office</span><span class="sxs-lookup"><span data-stu-id="2d3b4-103">Office</span></span>

<span data-ttu-id="2d3b4-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2d3b4-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="2d3b4-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="2d3b4-106">Requirements</span></span>

|<span data-ttu-id="2d3b4-107">要件</span><span class="sxs-lookup"><span data-stu-id="2d3b4-107">Requirement</span></span>| <span data-ttu-id="2d3b4-108">値</span><span class="sxs-lookup"><span data-stu-id="2d3b4-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d3b4-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2d3b4-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2d3b4-110">1.1</span><span class="sxs-lookup"><span data-stu-id="2d3b4-110">1.1</span></span>|
|[<span data-ttu-id="2d3b4-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2d3b4-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2d3b4-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2d3b4-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="2d3b4-113">Properties</span><span class="sxs-lookup"><span data-stu-id="2d3b4-113">Properties</span></span>

| <span data-ttu-id="2d3b4-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="2d3b4-114">Property</span></span> | <span data-ttu-id="2d3b4-115">モード</span><span class="sxs-lookup"><span data-stu-id="2d3b4-115">Modes</span></span> | <span data-ttu-id="2d3b4-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="2d3b4-116">Return type</span></span> | <span data-ttu-id="2d3b4-117">最小値</span><span class="sxs-lookup"><span data-stu-id="2d3b4-117">Minimum</span></span><br><span data-ttu-id="2d3b4-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="2d3b4-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2d3b4-119">context</span><span class="sxs-lookup"><span data-stu-id="2d3b4-119">context</span></span>](office.context.md) | <span data-ttu-id="2d3b4-120">作成</span><span class="sxs-lookup"><span data-stu-id="2d3b4-120">Compose</span></span><br><span data-ttu-id="2d3b4-121">読み取り</span><span class="sxs-lookup"><span data-stu-id="2d3b4-121">Read</span></span> | [<span data-ttu-id="2d3b4-122">Context</span><span class="sxs-lookup"><span data-stu-id="2d3b4-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7) | [<span data-ttu-id="2d3b4-123">1.1</span><span class="sxs-lookup"><span data-stu-id="2d3b4-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="2d3b4-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="2d3b4-124">Enumerations</span></span>

| <span data-ttu-id="2d3b4-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="2d3b4-125">Enumeration</span></span> | <span data-ttu-id="2d3b4-126">モード</span><span class="sxs-lookup"><span data-stu-id="2d3b4-126">Modes</span></span> | <span data-ttu-id="2d3b4-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="2d3b4-127">Return type</span></span> | <span data-ttu-id="2d3b4-128">最小値</span><span class="sxs-lookup"><span data-stu-id="2d3b4-128">Minimum</span></span><br><span data-ttu-id="2d3b4-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="2d3b4-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2d3b4-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="2d3b4-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="2d3b4-131">作成</span><span class="sxs-lookup"><span data-stu-id="2d3b4-131">Compose</span></span><br><span data-ttu-id="2d3b4-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="2d3b4-132">Read</span></span> | <span data-ttu-id="2d3b4-133">文字列</span><span class="sxs-lookup"><span data-stu-id="2d3b4-133">String</span></span> | [<span data-ttu-id="2d3b4-134">1.1</span><span class="sxs-lookup"><span data-stu-id="2d3b4-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2d3b4-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="2d3b4-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="2d3b4-136">作成</span><span class="sxs-lookup"><span data-stu-id="2d3b4-136">Compose</span></span><br><span data-ttu-id="2d3b4-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="2d3b4-137">Read</span></span> | <span data-ttu-id="2d3b4-138">文字列</span><span class="sxs-lookup"><span data-stu-id="2d3b4-138">String</span></span> | [<span data-ttu-id="2d3b4-139">1.1</span><span class="sxs-lookup"><span data-stu-id="2d3b4-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2d3b4-140">EventType</span><span class="sxs-lookup"><span data-stu-id="2d3b4-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="2d3b4-141">作成</span><span class="sxs-lookup"><span data-stu-id="2d3b4-141">Compose</span></span><br><span data-ttu-id="2d3b4-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="2d3b4-142">Read</span></span> | <span data-ttu-id="2d3b4-143">文字列</span><span class="sxs-lookup"><span data-stu-id="2d3b4-143">String</span></span> | [<span data-ttu-id="2d3b4-144">1.5</span><span class="sxs-lookup"><span data-stu-id="2d3b4-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="2d3b4-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="2d3b4-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="2d3b4-146">作成</span><span class="sxs-lookup"><span data-stu-id="2d3b4-146">Compose</span></span><br><span data-ttu-id="2d3b4-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="2d3b4-147">Read</span></span> | <span data-ttu-id="2d3b4-148">文字列</span><span class="sxs-lookup"><span data-stu-id="2d3b4-148">String</span></span> | [<span data-ttu-id="2d3b4-149">1.1</span><span class="sxs-lookup"><span data-stu-id="2d3b4-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="2d3b4-150">名前空間</span><span class="sxs-lookup"><span data-stu-id="2d3b4-150">Namespaces</span></span>

<span data-ttu-id="2d3b4-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="2d3b4-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="2d3b4-152">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="2d3b4-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="2d3b4-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="2d3b4-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="2d3b4-154">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="2d3b4-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="2d3b4-155">型</span><span class="sxs-lookup"><span data-stu-id="2d3b4-155">Type</span></span>

*   <span data-ttu-id="2d3b4-156">String</span><span class="sxs-lookup"><span data-stu-id="2d3b4-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2d3b4-157">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="2d3b4-157">Properties:</span></span>

|<span data-ttu-id="2d3b4-158">名前</span><span class="sxs-lookup"><span data-stu-id="2d3b4-158">Name</span></span>| <span data-ttu-id="2d3b4-159">種類</span><span class="sxs-lookup"><span data-stu-id="2d3b4-159">Type</span></span>| <span data-ttu-id="2d3b4-160">説明</span><span class="sxs-lookup"><span data-stu-id="2d3b4-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="2d3b4-161">文字列</span><span class="sxs-lookup"><span data-stu-id="2d3b4-161">String</span></span>|<span data-ttu-id="2d3b4-162">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="2d3b4-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="2d3b4-163">文字列</span><span class="sxs-lookup"><span data-stu-id="2d3b4-163">String</span></span>|<span data-ttu-id="2d3b4-164">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="2d3b4-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2d3b4-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="2d3b4-165">Requirements</span></span>

|<span data-ttu-id="2d3b4-166">要件</span><span class="sxs-lookup"><span data-stu-id="2d3b4-166">Requirement</span></span>| <span data-ttu-id="2d3b4-167">値</span><span class="sxs-lookup"><span data-stu-id="2d3b4-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d3b4-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2d3b4-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2d3b4-169">1.1</span><span class="sxs-lookup"><span data-stu-id="2d3b4-169">1.1</span></span>|
|[<span data-ttu-id="2d3b4-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2d3b4-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2d3b4-171">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2d3b4-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="2d3b4-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="2d3b4-172">CoercionType: String</span></span>

<span data-ttu-id="2d3b4-173">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="2d3b4-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2d3b4-174">型</span><span class="sxs-lookup"><span data-stu-id="2d3b4-174">Type</span></span>

*   <span data-ttu-id="2d3b4-175">String</span><span class="sxs-lookup"><span data-stu-id="2d3b4-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2d3b4-176">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="2d3b4-176">Properties:</span></span>

|<span data-ttu-id="2d3b4-177">名前</span><span class="sxs-lookup"><span data-stu-id="2d3b4-177">Name</span></span>| <span data-ttu-id="2d3b4-178">種類</span><span class="sxs-lookup"><span data-stu-id="2d3b4-178">Type</span></span>| <span data-ttu-id="2d3b4-179">説明</span><span class="sxs-lookup"><span data-stu-id="2d3b4-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="2d3b4-180">文字列</span><span class="sxs-lookup"><span data-stu-id="2d3b4-180">String</span></span>|<span data-ttu-id="2d3b4-181">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="2d3b4-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="2d3b4-182">String</span><span class="sxs-lookup"><span data-stu-id="2d3b4-182">String</span></span>|<span data-ttu-id="2d3b4-183">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="2d3b4-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2d3b4-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="2d3b4-184">Requirements</span></span>

|<span data-ttu-id="2d3b4-185">要件</span><span class="sxs-lookup"><span data-stu-id="2d3b4-185">Requirement</span></span>| <span data-ttu-id="2d3b4-186">値</span><span class="sxs-lookup"><span data-stu-id="2d3b4-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d3b4-187">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2d3b4-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2d3b4-188">1.1</span><span class="sxs-lookup"><span data-stu-id="2d3b4-188">1.1</span></span>|
|[<span data-ttu-id="2d3b4-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2d3b4-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2d3b4-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2d3b4-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="2d3b4-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="2d3b4-191">EventType: String</span></span>

<span data-ttu-id="2d3b4-192">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="2d3b4-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="2d3b4-193">型</span><span class="sxs-lookup"><span data-stu-id="2d3b4-193">Type</span></span>

*   <span data-ttu-id="2d3b4-194">String</span><span class="sxs-lookup"><span data-stu-id="2d3b4-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2d3b4-195">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="2d3b4-195">Properties:</span></span>

| <span data-ttu-id="2d3b4-196">名前</span><span class="sxs-lookup"><span data-stu-id="2d3b4-196">Name</span></span> | <span data-ttu-id="2d3b4-197">種類</span><span class="sxs-lookup"><span data-stu-id="2d3b4-197">Type</span></span> | <span data-ttu-id="2d3b4-198">説明</span><span class="sxs-lookup"><span data-stu-id="2d3b4-198">Description</span></span> | <span data-ttu-id="2d3b4-199">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="2d3b4-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="2d3b4-200">文字列</span><span class="sxs-lookup"><span data-stu-id="2d3b4-200">String</span></span> | <span data-ttu-id="2d3b4-201">選択した予定またはデータ系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="2d3b4-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="2d3b4-202">1.7</span><span class="sxs-lookup"><span data-stu-id="2d3b4-202">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="2d3b4-203">文字列</span><span class="sxs-lookup"><span data-stu-id="2d3b4-203">String</span></span> | <span data-ttu-id="2d3b4-204">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="2d3b4-204">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="2d3b4-205">1.5</span><span class="sxs-lookup"><span data-stu-id="2d3b4-205">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="2d3b4-206">文字列</span><span class="sxs-lookup"><span data-stu-id="2d3b4-206">String</span></span> | <span data-ttu-id="2d3b4-207">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="2d3b4-207">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="2d3b4-208">1.7</span><span class="sxs-lookup"><span data-stu-id="2d3b4-208">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="2d3b4-209">文字列</span><span class="sxs-lookup"><span data-stu-id="2d3b4-209">String</span></span> | <span data-ttu-id="2d3b4-210">選択したアイテムの定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="2d3b4-210">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="2d3b4-211">1.7</span><span class="sxs-lookup"><span data-stu-id="2d3b4-211">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2d3b4-212">Requirements</span><span class="sxs-lookup"><span data-stu-id="2d3b4-212">Requirements</span></span>

|<span data-ttu-id="2d3b4-213">要件</span><span class="sxs-lookup"><span data-stu-id="2d3b4-213">Requirement</span></span>| <span data-ttu-id="2d3b4-214">値</span><span class="sxs-lookup"><span data-stu-id="2d3b4-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d3b4-215">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2d3b4-215">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2d3b4-216">1.5</span><span class="sxs-lookup"><span data-stu-id="2d3b4-216">1.5</span></span> |
|[<span data-ttu-id="2d3b4-217">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2d3b4-217">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2d3b4-218">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2d3b4-218">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="2d3b4-219">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="2d3b4-219">SourceProperty: String</span></span>

<span data-ttu-id="2d3b4-220">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="2d3b4-220">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2d3b4-221">型</span><span class="sxs-lookup"><span data-stu-id="2d3b4-221">Type</span></span>

*   <span data-ttu-id="2d3b4-222">String</span><span class="sxs-lookup"><span data-stu-id="2d3b4-222">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2d3b4-223">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="2d3b4-223">Properties:</span></span>

|<span data-ttu-id="2d3b4-224">名前</span><span class="sxs-lookup"><span data-stu-id="2d3b4-224">Name</span></span>| <span data-ttu-id="2d3b4-225">種類</span><span class="sxs-lookup"><span data-stu-id="2d3b4-225">Type</span></span>| <span data-ttu-id="2d3b4-226">説明</span><span class="sxs-lookup"><span data-stu-id="2d3b4-226">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="2d3b4-227">文字列</span><span class="sxs-lookup"><span data-stu-id="2d3b4-227">String</span></span>|<span data-ttu-id="2d3b4-228">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="2d3b4-228">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="2d3b4-229">文字列</span><span class="sxs-lookup"><span data-stu-id="2d3b4-229">String</span></span>|<span data-ttu-id="2d3b4-230">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="2d3b4-230">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2d3b4-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="2d3b4-231">Requirements</span></span>

|<span data-ttu-id="2d3b4-232">要件</span><span class="sxs-lookup"><span data-stu-id="2d3b4-232">Requirement</span></span>| <span data-ttu-id="2d3b4-233">値</span><span class="sxs-lookup"><span data-stu-id="2d3b4-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d3b4-234">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2d3b4-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2d3b4-235">1.1</span><span class="sxs-lookup"><span data-stu-id="2d3b4-235">1.1</span></span>|
|[<span data-ttu-id="2d3b4-236">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2d3b4-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2d3b4-237">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2d3b4-237">Compose or Read</span></span>|
