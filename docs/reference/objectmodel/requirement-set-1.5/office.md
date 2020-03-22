---
title: Office 名前空間-要件セット1.5
description: メールボックス API 要件セット1.5 を使用した Outlook アドインで使用可能な Office 名前空間メンバー。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 848aa30c07b936c8454b2833d5dce3e1d15ee193
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891349"
---
# <a name="office-mailbox-requirement-set-15"></a><span data-ttu-id="2c13e-103">Office (メールボックス要件セット 1.5)</span><span class="sxs-lookup"><span data-stu-id="2c13e-103">Office (Mailbox requirement set 1.5)</span></span>

<span data-ttu-id="2c13e-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2c13e-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="2c13e-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="2c13e-106">Requirements</span></span>

|<span data-ttu-id="2c13e-107">要件</span><span class="sxs-lookup"><span data-stu-id="2c13e-107">Requirement</span></span>| <span data-ttu-id="2c13e-108">値</span><span class="sxs-lookup"><span data-stu-id="2c13e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c13e-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2c13e-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2c13e-110">1.1</span><span class="sxs-lookup"><span data-stu-id="2c13e-110">1.1</span></span>|
|[<span data-ttu-id="2c13e-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2c13e-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2c13e-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2c13e-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="2c13e-113">Properties</span><span class="sxs-lookup"><span data-stu-id="2c13e-113">Properties</span></span>

| <span data-ttu-id="2c13e-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="2c13e-114">Property</span></span> | <span data-ttu-id="2c13e-115">モード</span><span class="sxs-lookup"><span data-stu-id="2c13e-115">Modes</span></span> | <span data-ttu-id="2c13e-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="2c13e-116">Return type</span></span> | <span data-ttu-id="2c13e-117">最小値</span><span class="sxs-lookup"><span data-stu-id="2c13e-117">Minimum</span></span><br><span data-ttu-id="2c13e-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="2c13e-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2c13e-119">context</span><span class="sxs-lookup"><span data-stu-id="2c13e-119">context</span></span>](office.context.md) | <span data-ttu-id="2c13e-120">作成</span><span class="sxs-lookup"><span data-stu-id="2c13e-120">Compose</span></span><br><span data-ttu-id="2c13e-121">読み取り</span><span class="sxs-lookup"><span data-stu-id="2c13e-121">Read</span></span> | [<span data-ttu-id="2c13e-122">Context</span><span class="sxs-lookup"><span data-stu-id="2c13e-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.5) | [<span data-ttu-id="2c13e-123">1.1</span><span class="sxs-lookup"><span data-stu-id="2c13e-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="2c13e-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="2c13e-124">Enumerations</span></span>

| <span data-ttu-id="2c13e-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="2c13e-125">Enumeration</span></span> | <span data-ttu-id="2c13e-126">モード</span><span class="sxs-lookup"><span data-stu-id="2c13e-126">Modes</span></span> | <span data-ttu-id="2c13e-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="2c13e-127">Return type</span></span> | <span data-ttu-id="2c13e-128">最小値</span><span class="sxs-lookup"><span data-stu-id="2c13e-128">Minimum</span></span><br><span data-ttu-id="2c13e-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="2c13e-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2c13e-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="2c13e-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="2c13e-131">作成</span><span class="sxs-lookup"><span data-stu-id="2c13e-131">Compose</span></span><br><span data-ttu-id="2c13e-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="2c13e-132">Read</span></span> | <span data-ttu-id="2c13e-133">String</span><span class="sxs-lookup"><span data-stu-id="2c13e-133">String</span></span> | [<span data-ttu-id="2c13e-134">1.1</span><span class="sxs-lookup"><span data-stu-id="2c13e-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2c13e-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="2c13e-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="2c13e-136">作成</span><span class="sxs-lookup"><span data-stu-id="2c13e-136">Compose</span></span><br><span data-ttu-id="2c13e-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="2c13e-137">Read</span></span> | <span data-ttu-id="2c13e-138">String</span><span class="sxs-lookup"><span data-stu-id="2c13e-138">String</span></span> | [<span data-ttu-id="2c13e-139">1.1</span><span class="sxs-lookup"><span data-stu-id="2c13e-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2c13e-140">EventType</span><span class="sxs-lookup"><span data-stu-id="2c13e-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="2c13e-141">作成</span><span class="sxs-lookup"><span data-stu-id="2c13e-141">Compose</span></span><br><span data-ttu-id="2c13e-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="2c13e-142">Read</span></span> | <span data-ttu-id="2c13e-143">String</span><span class="sxs-lookup"><span data-stu-id="2c13e-143">String</span></span> | [<span data-ttu-id="2c13e-144">1.5</span><span class="sxs-lookup"><span data-stu-id="2c13e-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="2c13e-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="2c13e-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="2c13e-146">作成</span><span class="sxs-lookup"><span data-stu-id="2c13e-146">Compose</span></span><br><span data-ttu-id="2c13e-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="2c13e-147">Read</span></span> | <span data-ttu-id="2c13e-148">String</span><span class="sxs-lookup"><span data-stu-id="2c13e-148">String</span></span> | [<span data-ttu-id="2c13e-149">1.1</span><span class="sxs-lookup"><span data-stu-id="2c13e-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="2c13e-150">名前空間</span><span class="sxs-lookup"><span data-stu-id="2c13e-150">Namespaces</span></span>

<span data-ttu-id="2c13e-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="2c13e-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="2c13e-152">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="2c13e-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="2c13e-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="2c13e-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="2c13e-154">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="2c13e-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="2c13e-155">型</span><span class="sxs-lookup"><span data-stu-id="2c13e-155">Type</span></span>

*   <span data-ttu-id="2c13e-156">String</span><span class="sxs-lookup"><span data-stu-id="2c13e-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2c13e-157">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="2c13e-157">Properties:</span></span>

|<span data-ttu-id="2c13e-158">名前</span><span class="sxs-lookup"><span data-stu-id="2c13e-158">Name</span></span>| <span data-ttu-id="2c13e-159">種類</span><span class="sxs-lookup"><span data-stu-id="2c13e-159">Type</span></span>| <span data-ttu-id="2c13e-160">説明</span><span class="sxs-lookup"><span data-stu-id="2c13e-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="2c13e-161">String</span><span class="sxs-lookup"><span data-stu-id="2c13e-161">String</span></span>|<span data-ttu-id="2c13e-162">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="2c13e-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="2c13e-163">String</span><span class="sxs-lookup"><span data-stu-id="2c13e-163">String</span></span>|<span data-ttu-id="2c13e-164">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="2c13e-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2c13e-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="2c13e-165">Requirements</span></span>

|<span data-ttu-id="2c13e-166">要件</span><span class="sxs-lookup"><span data-stu-id="2c13e-166">Requirement</span></span>| <span data-ttu-id="2c13e-167">値</span><span class="sxs-lookup"><span data-stu-id="2c13e-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c13e-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2c13e-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2c13e-169">1.1</span><span class="sxs-lookup"><span data-stu-id="2c13e-169">1.1</span></span>|
|[<span data-ttu-id="2c13e-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2c13e-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2c13e-171">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2c13e-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="2c13e-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="2c13e-172">CoercionType: String</span></span>

<span data-ttu-id="2c13e-173">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="2c13e-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2c13e-174">型</span><span class="sxs-lookup"><span data-stu-id="2c13e-174">Type</span></span>

*   <span data-ttu-id="2c13e-175">String</span><span class="sxs-lookup"><span data-stu-id="2c13e-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2c13e-176">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="2c13e-176">Properties:</span></span>

|<span data-ttu-id="2c13e-177">名前</span><span class="sxs-lookup"><span data-stu-id="2c13e-177">Name</span></span>| <span data-ttu-id="2c13e-178">種類</span><span class="sxs-lookup"><span data-stu-id="2c13e-178">Type</span></span>| <span data-ttu-id="2c13e-179">説明</span><span class="sxs-lookup"><span data-stu-id="2c13e-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="2c13e-180">String</span><span class="sxs-lookup"><span data-stu-id="2c13e-180">String</span></span>|<span data-ttu-id="2c13e-181">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="2c13e-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="2c13e-182">String</span><span class="sxs-lookup"><span data-stu-id="2c13e-182">String</span></span>|<span data-ttu-id="2c13e-183">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="2c13e-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2c13e-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="2c13e-184">Requirements</span></span>

|<span data-ttu-id="2c13e-185">要件</span><span class="sxs-lookup"><span data-stu-id="2c13e-185">Requirement</span></span>| <span data-ttu-id="2c13e-186">値</span><span class="sxs-lookup"><span data-stu-id="2c13e-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c13e-187">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2c13e-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2c13e-188">1.1</span><span class="sxs-lookup"><span data-stu-id="2c13e-188">1.1</span></span>|
|[<span data-ttu-id="2c13e-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2c13e-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2c13e-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2c13e-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="2c13e-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="2c13e-191">EventType: String</span></span>

<span data-ttu-id="2c13e-192">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="2c13e-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="2c13e-193">型</span><span class="sxs-lookup"><span data-stu-id="2c13e-193">Type</span></span>

*   <span data-ttu-id="2c13e-194">String</span><span class="sxs-lookup"><span data-stu-id="2c13e-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2c13e-195">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="2c13e-195">Properties:</span></span>

| <span data-ttu-id="2c13e-196">名前</span><span class="sxs-lookup"><span data-stu-id="2c13e-196">Name</span></span> | <span data-ttu-id="2c13e-197">種類</span><span class="sxs-lookup"><span data-stu-id="2c13e-197">Type</span></span> | <span data-ttu-id="2c13e-198">説明</span><span class="sxs-lookup"><span data-stu-id="2c13e-198">Description</span></span> | <span data-ttu-id="2c13e-199">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="2c13e-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="2c13e-200">String</span><span class="sxs-lookup"><span data-stu-id="2c13e-200">String</span></span> | <span data-ttu-id="2c13e-201">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="2c13e-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="2c13e-202">1.5</span><span class="sxs-lookup"><span data-stu-id="2c13e-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2c13e-203">Requirements</span><span class="sxs-lookup"><span data-stu-id="2c13e-203">Requirements</span></span>

|<span data-ttu-id="2c13e-204">要件</span><span class="sxs-lookup"><span data-stu-id="2c13e-204">Requirement</span></span>| <span data-ttu-id="2c13e-205">値</span><span class="sxs-lookup"><span data-stu-id="2c13e-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c13e-206">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2c13e-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2c13e-207">1.5</span><span class="sxs-lookup"><span data-stu-id="2c13e-207">1.5</span></span> |
|[<span data-ttu-id="2c13e-208">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2c13e-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2c13e-209">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2c13e-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="2c13e-210">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="2c13e-210">SourceProperty: String</span></span>

<span data-ttu-id="2c13e-211">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="2c13e-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2c13e-212">型</span><span class="sxs-lookup"><span data-stu-id="2c13e-212">Type</span></span>

*   <span data-ttu-id="2c13e-213">String</span><span class="sxs-lookup"><span data-stu-id="2c13e-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2c13e-214">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="2c13e-214">Properties:</span></span>

|<span data-ttu-id="2c13e-215">名前</span><span class="sxs-lookup"><span data-stu-id="2c13e-215">Name</span></span>| <span data-ttu-id="2c13e-216">種類</span><span class="sxs-lookup"><span data-stu-id="2c13e-216">Type</span></span>| <span data-ttu-id="2c13e-217">説明</span><span class="sxs-lookup"><span data-stu-id="2c13e-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="2c13e-218">String</span><span class="sxs-lookup"><span data-stu-id="2c13e-218">String</span></span>|<span data-ttu-id="2c13e-219">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="2c13e-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="2c13e-220">String</span><span class="sxs-lookup"><span data-stu-id="2c13e-220">String</span></span>|<span data-ttu-id="2c13e-221">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="2c13e-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2c13e-222">Requirements</span><span class="sxs-lookup"><span data-stu-id="2c13e-222">Requirements</span></span>

|<span data-ttu-id="2c13e-223">要件</span><span class="sxs-lookup"><span data-stu-id="2c13e-223">Requirement</span></span>| <span data-ttu-id="2c13e-224">値</span><span class="sxs-lookup"><span data-stu-id="2c13e-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c13e-225">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2c13e-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2c13e-226">1.1</span><span class="sxs-lookup"><span data-stu-id="2c13e-226">1.1</span></span>|
|[<span data-ttu-id="2c13e-227">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2c13e-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2c13e-228">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2c13e-228">Compose or Read</span></span>|
