---
title: Office 名前空間-要件セット1.6
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: e15f01db9423a9df38608f18098d2c808f5d944b
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814669"
---
# <a name="office"></a><span data-ttu-id="99c20-102">Office</span><span class="sxs-lookup"><span data-stu-id="99c20-102">Office</span></span>

<span data-ttu-id="99c20-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="99c20-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="99c20-105">要件</span><span class="sxs-lookup"><span data-stu-id="99c20-105">Requirements</span></span>

|<span data-ttu-id="99c20-106">要件</span><span class="sxs-lookup"><span data-stu-id="99c20-106">Requirement</span></span>| <span data-ttu-id="99c20-107">値</span><span class="sxs-lookup"><span data-stu-id="99c20-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="99c20-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="99c20-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="99c20-109">1.1</span><span class="sxs-lookup"><span data-stu-id="99c20-109">1.1</span></span>|
|[<span data-ttu-id="99c20-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="99c20-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99c20-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="99c20-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="99c20-112">Properties</span><span class="sxs-lookup"><span data-stu-id="99c20-112">Properties</span></span>

| <span data-ttu-id="99c20-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="99c20-113">Property</span></span> | <span data-ttu-id="99c20-114">モード</span><span class="sxs-lookup"><span data-stu-id="99c20-114">Modes</span></span> | <span data-ttu-id="99c20-115">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="99c20-115">Return type</span></span> | <span data-ttu-id="99c20-116">最小値</span><span class="sxs-lookup"><span data-stu-id="99c20-116">Minimum</span></span><br><span data-ttu-id="99c20-117">要件セット</span><span class="sxs-lookup"><span data-stu-id="99c20-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="99c20-118">context</span><span class="sxs-lookup"><span data-stu-id="99c20-118">context</span></span>](office.context.md) | <span data-ttu-id="99c20-119">作成</span><span class="sxs-lookup"><span data-stu-id="99c20-119">Compose</span></span><br><span data-ttu-id="99c20-120">読み取り</span><span class="sxs-lookup"><span data-stu-id="99c20-120">Read</span></span> | [<span data-ttu-id="99c20-121">Context</span><span class="sxs-lookup"><span data-stu-id="99c20-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6) | [<span data-ttu-id="99c20-122">1.1</span><span class="sxs-lookup"><span data-stu-id="99c20-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="99c20-123">列挙型</span><span class="sxs-lookup"><span data-stu-id="99c20-123">Enumerations</span></span>

| <span data-ttu-id="99c20-124">列挙体</span><span class="sxs-lookup"><span data-stu-id="99c20-124">Enumeration</span></span> | <span data-ttu-id="99c20-125">モード</span><span class="sxs-lookup"><span data-stu-id="99c20-125">Modes</span></span> | <span data-ttu-id="99c20-126">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="99c20-126">Return type</span></span> | <span data-ttu-id="99c20-127">最小値</span><span class="sxs-lookup"><span data-stu-id="99c20-127">Minimum</span></span><br><span data-ttu-id="99c20-128">要件セット</span><span class="sxs-lookup"><span data-stu-id="99c20-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="99c20-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="99c20-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="99c20-130">作成</span><span class="sxs-lookup"><span data-stu-id="99c20-130">Compose</span></span><br><span data-ttu-id="99c20-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="99c20-131">Read</span></span> | <span data-ttu-id="99c20-132">String</span><span class="sxs-lookup"><span data-stu-id="99c20-132">String</span></span> | [<span data-ttu-id="99c20-133">1.1</span><span class="sxs-lookup"><span data-stu-id="99c20-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="99c20-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="99c20-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="99c20-135">作成</span><span class="sxs-lookup"><span data-stu-id="99c20-135">Compose</span></span><br><span data-ttu-id="99c20-136">読み取り</span><span class="sxs-lookup"><span data-stu-id="99c20-136">Read</span></span> | <span data-ttu-id="99c20-137">String</span><span class="sxs-lookup"><span data-stu-id="99c20-137">String</span></span> | [<span data-ttu-id="99c20-138">1.1</span><span class="sxs-lookup"><span data-stu-id="99c20-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="99c20-139">EventType</span><span class="sxs-lookup"><span data-stu-id="99c20-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="99c20-140">作成</span><span class="sxs-lookup"><span data-stu-id="99c20-140">Compose</span></span><br><span data-ttu-id="99c20-141">読み取り</span><span class="sxs-lookup"><span data-stu-id="99c20-141">Read</span></span> | <span data-ttu-id="99c20-142">String</span><span class="sxs-lookup"><span data-stu-id="99c20-142">String</span></span> | [<span data-ttu-id="99c20-143">1.5</span><span class="sxs-lookup"><span data-stu-id="99c20-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="99c20-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="99c20-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="99c20-145">作成</span><span class="sxs-lookup"><span data-stu-id="99c20-145">Compose</span></span><br><span data-ttu-id="99c20-146">読み取り</span><span class="sxs-lookup"><span data-stu-id="99c20-146">Read</span></span> | <span data-ttu-id="99c20-147">String</span><span class="sxs-lookup"><span data-stu-id="99c20-147">String</span></span> | [<span data-ttu-id="99c20-148">1.1</span><span class="sxs-lookup"><span data-stu-id="99c20-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="99c20-149">名前空間</span><span class="sxs-lookup"><span data-stu-id="99c20-149">Namespaces</span></span>

<span data-ttu-id="99c20-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="99c20-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="99c20-151">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="99c20-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="99c20-152">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="99c20-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="99c20-153">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="99c20-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="99c20-154">型</span><span class="sxs-lookup"><span data-stu-id="99c20-154">Type</span></span>

*   <span data-ttu-id="99c20-155">String</span><span class="sxs-lookup"><span data-stu-id="99c20-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="99c20-156">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="99c20-156">Properties:</span></span>

|<span data-ttu-id="99c20-157">名前</span><span class="sxs-lookup"><span data-stu-id="99c20-157">Name</span></span>| <span data-ttu-id="99c20-158">種類</span><span class="sxs-lookup"><span data-stu-id="99c20-158">Type</span></span>| <span data-ttu-id="99c20-159">説明</span><span class="sxs-lookup"><span data-stu-id="99c20-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="99c20-160">String</span><span class="sxs-lookup"><span data-stu-id="99c20-160">String</span></span>|<span data-ttu-id="99c20-161">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="99c20-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="99c20-162">String</span><span class="sxs-lookup"><span data-stu-id="99c20-162">String</span></span>|<span data-ttu-id="99c20-163">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="99c20-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="99c20-164">要件</span><span class="sxs-lookup"><span data-stu-id="99c20-164">Requirements</span></span>

|<span data-ttu-id="99c20-165">要件</span><span class="sxs-lookup"><span data-stu-id="99c20-165">Requirement</span></span>| <span data-ttu-id="99c20-166">値</span><span class="sxs-lookup"><span data-stu-id="99c20-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="99c20-167">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="99c20-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="99c20-168">1.1</span><span class="sxs-lookup"><span data-stu-id="99c20-168">1.1</span></span>|
|[<span data-ttu-id="99c20-169">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="99c20-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99c20-170">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="99c20-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="99c20-171">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="99c20-171">CoercionType: String</span></span>

<span data-ttu-id="99c20-172">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="99c20-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="99c20-173">型</span><span class="sxs-lookup"><span data-stu-id="99c20-173">Type</span></span>

*   <span data-ttu-id="99c20-174">String</span><span class="sxs-lookup"><span data-stu-id="99c20-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="99c20-175">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="99c20-175">Properties:</span></span>

|<span data-ttu-id="99c20-176">名前</span><span class="sxs-lookup"><span data-stu-id="99c20-176">Name</span></span>| <span data-ttu-id="99c20-177">種類</span><span class="sxs-lookup"><span data-stu-id="99c20-177">Type</span></span>| <span data-ttu-id="99c20-178">説明</span><span class="sxs-lookup"><span data-stu-id="99c20-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="99c20-179">String</span><span class="sxs-lookup"><span data-stu-id="99c20-179">String</span></span>|<span data-ttu-id="99c20-180">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="99c20-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="99c20-181">String</span><span class="sxs-lookup"><span data-stu-id="99c20-181">String</span></span>|<span data-ttu-id="99c20-182">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="99c20-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="99c20-183">要件</span><span class="sxs-lookup"><span data-stu-id="99c20-183">Requirements</span></span>

|<span data-ttu-id="99c20-184">要件</span><span class="sxs-lookup"><span data-stu-id="99c20-184">Requirement</span></span>| <span data-ttu-id="99c20-185">値</span><span class="sxs-lookup"><span data-stu-id="99c20-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="99c20-186">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="99c20-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="99c20-187">1.1</span><span class="sxs-lookup"><span data-stu-id="99c20-187">1.1</span></span>|
|[<span data-ttu-id="99c20-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="99c20-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99c20-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="99c20-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="99c20-190">EventType: String</span><span class="sxs-lookup"><span data-stu-id="99c20-190">EventType: String</span></span>

<span data-ttu-id="99c20-191">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="99c20-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="99c20-192">型</span><span class="sxs-lookup"><span data-stu-id="99c20-192">Type</span></span>

*   <span data-ttu-id="99c20-193">String</span><span class="sxs-lookup"><span data-stu-id="99c20-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="99c20-194">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="99c20-194">Properties:</span></span>

| <span data-ttu-id="99c20-195">名前</span><span class="sxs-lookup"><span data-stu-id="99c20-195">Name</span></span> | <span data-ttu-id="99c20-196">種類</span><span class="sxs-lookup"><span data-stu-id="99c20-196">Type</span></span> | <span data-ttu-id="99c20-197">説明</span><span class="sxs-lookup"><span data-stu-id="99c20-197">Description</span></span> | <span data-ttu-id="99c20-198">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="99c20-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="99c20-199">String</span><span class="sxs-lookup"><span data-stu-id="99c20-199">String</span></span> | <span data-ttu-id="99c20-200">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="99c20-200">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="99c20-201">1.5</span><span class="sxs-lookup"><span data-stu-id="99c20-201">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="99c20-202">要件</span><span class="sxs-lookup"><span data-stu-id="99c20-202">Requirements</span></span>

|<span data-ttu-id="99c20-203">要件</span><span class="sxs-lookup"><span data-stu-id="99c20-203">Requirement</span></span>| <span data-ttu-id="99c20-204">値</span><span class="sxs-lookup"><span data-stu-id="99c20-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="99c20-205">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="99c20-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="99c20-206">1.5</span><span class="sxs-lookup"><span data-stu-id="99c20-206">1.5</span></span> |
|[<span data-ttu-id="99c20-207">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="99c20-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99c20-208">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="99c20-208">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="99c20-209">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="99c20-209">SourceProperty: String</span></span>

<span data-ttu-id="99c20-210">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="99c20-210">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="99c20-211">型</span><span class="sxs-lookup"><span data-stu-id="99c20-211">Type</span></span>

*   <span data-ttu-id="99c20-212">String</span><span class="sxs-lookup"><span data-stu-id="99c20-212">String</span></span>

##### <a name="properties"></a><span data-ttu-id="99c20-213">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="99c20-213">Properties:</span></span>

|<span data-ttu-id="99c20-214">名前</span><span class="sxs-lookup"><span data-stu-id="99c20-214">Name</span></span>| <span data-ttu-id="99c20-215">種類</span><span class="sxs-lookup"><span data-stu-id="99c20-215">Type</span></span>| <span data-ttu-id="99c20-216">説明</span><span class="sxs-lookup"><span data-stu-id="99c20-216">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="99c20-217">String</span><span class="sxs-lookup"><span data-stu-id="99c20-217">String</span></span>|<span data-ttu-id="99c20-218">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="99c20-218">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="99c20-219">String</span><span class="sxs-lookup"><span data-stu-id="99c20-219">String</span></span>|<span data-ttu-id="99c20-220">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="99c20-220">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="99c20-221">要件</span><span class="sxs-lookup"><span data-stu-id="99c20-221">Requirements</span></span>

|<span data-ttu-id="99c20-222">要件</span><span class="sxs-lookup"><span data-stu-id="99c20-222">Requirement</span></span>| <span data-ttu-id="99c20-223">値</span><span class="sxs-lookup"><span data-stu-id="99c20-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="99c20-224">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="99c20-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="99c20-225">1.1</span><span class="sxs-lookup"><span data-stu-id="99c20-225">1.1</span></span>|
|[<span data-ttu-id="99c20-226">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="99c20-226">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="99c20-227">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="99c20-227">Compose or Read</span></span>|
