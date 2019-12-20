---
title: Office 名前空間-要件セット1.5
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 63dbb3ac10492ac6e2019353b8cb057227e4c1e6
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814753"
---
# <a name="office"></a><span data-ttu-id="1f7de-102">Office</span><span class="sxs-lookup"><span data-stu-id="1f7de-102">Office</span></span>

<span data-ttu-id="1f7de-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1f7de-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f7de-105">要件</span><span class="sxs-lookup"><span data-stu-id="1f7de-105">Requirements</span></span>

|<span data-ttu-id="1f7de-106">要件</span><span class="sxs-lookup"><span data-stu-id="1f7de-106">Requirement</span></span>| <span data-ttu-id="1f7de-107">値</span><span class="sxs-lookup"><span data-stu-id="1f7de-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f7de-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1f7de-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1f7de-109">1.1</span><span class="sxs-lookup"><span data-stu-id="1f7de-109">1.1</span></span>|
|[<span data-ttu-id="1f7de-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1f7de-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1f7de-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1f7de-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="1f7de-112">Properties</span><span class="sxs-lookup"><span data-stu-id="1f7de-112">Properties</span></span>

| <span data-ttu-id="1f7de-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="1f7de-113">Property</span></span> | <span data-ttu-id="1f7de-114">モード</span><span class="sxs-lookup"><span data-stu-id="1f7de-114">Modes</span></span> | <span data-ttu-id="1f7de-115">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="1f7de-115">Return type</span></span> | <span data-ttu-id="1f7de-116">最小値</span><span class="sxs-lookup"><span data-stu-id="1f7de-116">Minimum</span></span><br><span data-ttu-id="1f7de-117">要件セット</span><span class="sxs-lookup"><span data-stu-id="1f7de-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="1f7de-118">context</span><span class="sxs-lookup"><span data-stu-id="1f7de-118">context</span></span>](office.context.md) | <span data-ttu-id="1f7de-119">作成</span><span class="sxs-lookup"><span data-stu-id="1f7de-119">Compose</span></span><br><span data-ttu-id="1f7de-120">読み取り</span><span class="sxs-lookup"><span data-stu-id="1f7de-120">Read</span></span> | [<span data-ttu-id="1f7de-121">Context</span><span class="sxs-lookup"><span data-stu-id="1f7de-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.5) | [<span data-ttu-id="1f7de-122">1.1</span><span class="sxs-lookup"><span data-stu-id="1f7de-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="1f7de-123">列挙型</span><span class="sxs-lookup"><span data-stu-id="1f7de-123">Enumerations</span></span>

| <span data-ttu-id="1f7de-124">列挙体</span><span class="sxs-lookup"><span data-stu-id="1f7de-124">Enumeration</span></span> | <span data-ttu-id="1f7de-125">モード</span><span class="sxs-lookup"><span data-stu-id="1f7de-125">Modes</span></span> | <span data-ttu-id="1f7de-126">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="1f7de-126">Return type</span></span> | <span data-ttu-id="1f7de-127">最小値</span><span class="sxs-lookup"><span data-stu-id="1f7de-127">Minimum</span></span><br><span data-ttu-id="1f7de-128">要件セット</span><span class="sxs-lookup"><span data-stu-id="1f7de-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="1f7de-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="1f7de-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="1f7de-130">作成</span><span class="sxs-lookup"><span data-stu-id="1f7de-130">Compose</span></span><br><span data-ttu-id="1f7de-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="1f7de-131">Read</span></span> | <span data-ttu-id="1f7de-132">String</span><span class="sxs-lookup"><span data-stu-id="1f7de-132">String</span></span> | [<span data-ttu-id="1f7de-133">1.1</span><span class="sxs-lookup"><span data-stu-id="1f7de-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1f7de-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="1f7de-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="1f7de-135">作成</span><span class="sxs-lookup"><span data-stu-id="1f7de-135">Compose</span></span><br><span data-ttu-id="1f7de-136">読み取り</span><span class="sxs-lookup"><span data-stu-id="1f7de-136">Read</span></span> | <span data-ttu-id="1f7de-137">String</span><span class="sxs-lookup"><span data-stu-id="1f7de-137">String</span></span> | [<span data-ttu-id="1f7de-138">1.1</span><span class="sxs-lookup"><span data-stu-id="1f7de-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1f7de-139">EventType</span><span class="sxs-lookup"><span data-stu-id="1f7de-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="1f7de-140">作成</span><span class="sxs-lookup"><span data-stu-id="1f7de-140">Compose</span></span><br><span data-ttu-id="1f7de-141">読み取り</span><span class="sxs-lookup"><span data-stu-id="1f7de-141">Read</span></span> | <span data-ttu-id="1f7de-142">String</span><span class="sxs-lookup"><span data-stu-id="1f7de-142">String</span></span> | [<span data-ttu-id="1f7de-143">1.5</span><span class="sxs-lookup"><span data-stu-id="1f7de-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="1f7de-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="1f7de-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="1f7de-145">作成</span><span class="sxs-lookup"><span data-stu-id="1f7de-145">Compose</span></span><br><span data-ttu-id="1f7de-146">読み取り</span><span class="sxs-lookup"><span data-stu-id="1f7de-146">Read</span></span> | <span data-ttu-id="1f7de-147">String</span><span class="sxs-lookup"><span data-stu-id="1f7de-147">String</span></span> | [<span data-ttu-id="1f7de-148">1.1</span><span class="sxs-lookup"><span data-stu-id="1f7de-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="1f7de-149">名前空間</span><span class="sxs-lookup"><span data-stu-id="1f7de-149">Namespaces</span></span>

<span data-ttu-id="1f7de-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="1f7de-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="1f7de-151">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="1f7de-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="1f7de-152">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="1f7de-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="1f7de-153">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="1f7de-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="1f7de-154">型</span><span class="sxs-lookup"><span data-stu-id="1f7de-154">Type</span></span>

*   <span data-ttu-id="1f7de-155">String</span><span class="sxs-lookup"><span data-stu-id="1f7de-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1f7de-156">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="1f7de-156">Properties:</span></span>

|<span data-ttu-id="1f7de-157">名前</span><span class="sxs-lookup"><span data-stu-id="1f7de-157">Name</span></span>| <span data-ttu-id="1f7de-158">種類</span><span class="sxs-lookup"><span data-stu-id="1f7de-158">Type</span></span>| <span data-ttu-id="1f7de-159">説明</span><span class="sxs-lookup"><span data-stu-id="1f7de-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="1f7de-160">String</span><span class="sxs-lookup"><span data-stu-id="1f7de-160">String</span></span>|<span data-ttu-id="1f7de-161">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="1f7de-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="1f7de-162">String</span><span class="sxs-lookup"><span data-stu-id="1f7de-162">String</span></span>|<span data-ttu-id="1f7de-163">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="1f7de-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1f7de-164">要件</span><span class="sxs-lookup"><span data-stu-id="1f7de-164">Requirements</span></span>

|<span data-ttu-id="1f7de-165">要件</span><span class="sxs-lookup"><span data-stu-id="1f7de-165">Requirement</span></span>| <span data-ttu-id="1f7de-166">値</span><span class="sxs-lookup"><span data-stu-id="1f7de-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f7de-167">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1f7de-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1f7de-168">1.1</span><span class="sxs-lookup"><span data-stu-id="1f7de-168">1.1</span></span>|
|[<span data-ttu-id="1f7de-169">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1f7de-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1f7de-170">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1f7de-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="1f7de-171">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="1f7de-171">CoercionType: String</span></span>

<span data-ttu-id="1f7de-172">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="1f7de-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="1f7de-173">型</span><span class="sxs-lookup"><span data-stu-id="1f7de-173">Type</span></span>

*   <span data-ttu-id="1f7de-174">String</span><span class="sxs-lookup"><span data-stu-id="1f7de-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1f7de-175">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="1f7de-175">Properties:</span></span>

|<span data-ttu-id="1f7de-176">名前</span><span class="sxs-lookup"><span data-stu-id="1f7de-176">Name</span></span>| <span data-ttu-id="1f7de-177">種類</span><span class="sxs-lookup"><span data-stu-id="1f7de-177">Type</span></span>| <span data-ttu-id="1f7de-178">説明</span><span class="sxs-lookup"><span data-stu-id="1f7de-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="1f7de-179">String</span><span class="sxs-lookup"><span data-stu-id="1f7de-179">String</span></span>|<span data-ttu-id="1f7de-180">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="1f7de-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="1f7de-181">String</span><span class="sxs-lookup"><span data-stu-id="1f7de-181">String</span></span>|<span data-ttu-id="1f7de-182">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="1f7de-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1f7de-183">要件</span><span class="sxs-lookup"><span data-stu-id="1f7de-183">Requirements</span></span>

|<span data-ttu-id="1f7de-184">要件</span><span class="sxs-lookup"><span data-stu-id="1f7de-184">Requirement</span></span>| <span data-ttu-id="1f7de-185">値</span><span class="sxs-lookup"><span data-stu-id="1f7de-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f7de-186">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1f7de-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1f7de-187">1.1</span><span class="sxs-lookup"><span data-stu-id="1f7de-187">1.1</span></span>|
|[<span data-ttu-id="1f7de-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1f7de-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1f7de-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1f7de-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="1f7de-190">EventType: String</span><span class="sxs-lookup"><span data-stu-id="1f7de-190">EventType: String</span></span>

<span data-ttu-id="1f7de-191">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="1f7de-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="1f7de-192">型</span><span class="sxs-lookup"><span data-stu-id="1f7de-192">Type</span></span>

*   <span data-ttu-id="1f7de-193">String</span><span class="sxs-lookup"><span data-stu-id="1f7de-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1f7de-194">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="1f7de-194">Properties:</span></span>

| <span data-ttu-id="1f7de-195">名前</span><span class="sxs-lookup"><span data-stu-id="1f7de-195">Name</span></span> | <span data-ttu-id="1f7de-196">種類</span><span class="sxs-lookup"><span data-stu-id="1f7de-196">Type</span></span> | <span data-ttu-id="1f7de-197">説明</span><span class="sxs-lookup"><span data-stu-id="1f7de-197">Description</span></span> | <span data-ttu-id="1f7de-198">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="1f7de-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="1f7de-199">String</span><span class="sxs-lookup"><span data-stu-id="1f7de-199">String</span></span> | <span data-ttu-id="1f7de-200">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="1f7de-200">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="1f7de-201">1.5</span><span class="sxs-lookup"><span data-stu-id="1f7de-201">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1f7de-202">要件</span><span class="sxs-lookup"><span data-stu-id="1f7de-202">Requirements</span></span>

|<span data-ttu-id="1f7de-203">要件</span><span class="sxs-lookup"><span data-stu-id="1f7de-203">Requirement</span></span>| <span data-ttu-id="1f7de-204">値</span><span class="sxs-lookup"><span data-stu-id="1f7de-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f7de-205">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1f7de-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1f7de-206">1.5</span><span class="sxs-lookup"><span data-stu-id="1f7de-206">1.5</span></span> |
|[<span data-ttu-id="1f7de-207">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1f7de-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1f7de-208">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1f7de-208">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="1f7de-209">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="1f7de-209">SourceProperty: String</span></span>

<span data-ttu-id="1f7de-210">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="1f7de-210">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="1f7de-211">型</span><span class="sxs-lookup"><span data-stu-id="1f7de-211">Type</span></span>

*   <span data-ttu-id="1f7de-212">String</span><span class="sxs-lookup"><span data-stu-id="1f7de-212">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1f7de-213">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="1f7de-213">Properties:</span></span>

|<span data-ttu-id="1f7de-214">名前</span><span class="sxs-lookup"><span data-stu-id="1f7de-214">Name</span></span>| <span data-ttu-id="1f7de-215">種類</span><span class="sxs-lookup"><span data-stu-id="1f7de-215">Type</span></span>| <span data-ttu-id="1f7de-216">説明</span><span class="sxs-lookup"><span data-stu-id="1f7de-216">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="1f7de-217">String</span><span class="sxs-lookup"><span data-stu-id="1f7de-217">String</span></span>|<span data-ttu-id="1f7de-218">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="1f7de-218">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="1f7de-219">String</span><span class="sxs-lookup"><span data-stu-id="1f7de-219">String</span></span>|<span data-ttu-id="1f7de-220">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="1f7de-220">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1f7de-221">要件</span><span class="sxs-lookup"><span data-stu-id="1f7de-221">Requirements</span></span>

|<span data-ttu-id="1f7de-222">要件</span><span class="sxs-lookup"><span data-stu-id="1f7de-222">Requirement</span></span>| <span data-ttu-id="1f7de-223">値</span><span class="sxs-lookup"><span data-stu-id="1f7de-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f7de-224">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1f7de-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1f7de-225">1.1</span><span class="sxs-lookup"><span data-stu-id="1f7de-225">1.1</span></span>|
|[<span data-ttu-id="1f7de-226">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1f7de-226">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1f7de-227">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1f7de-227">Compose or Read</span></span>|
