---
title: Office 名前空間-要件セット1.5
description: Outlook アドイン API の最上位レベルの名前空間のオブジェクトモデル (Mailbox API 1.5 バージョン)。
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: ed65472de4acbe4f610e0355cc5de734938149ef
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720023"
---
# <a name="office"></a><span data-ttu-id="0266d-103">Office</span><span class="sxs-lookup"><span data-stu-id="0266d-103">Office</span></span>

<span data-ttu-id="0266d-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0266d-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="0266d-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="0266d-106">Requirements</span></span>

|<span data-ttu-id="0266d-107">要件</span><span class="sxs-lookup"><span data-stu-id="0266d-107">Requirement</span></span>| <span data-ttu-id="0266d-108">値</span><span class="sxs-lookup"><span data-stu-id="0266d-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="0266d-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0266d-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0266d-110">1.1</span><span class="sxs-lookup"><span data-stu-id="0266d-110">1.1</span></span>|
|[<span data-ttu-id="0266d-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0266d-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0266d-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0266d-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="0266d-113">Properties</span><span class="sxs-lookup"><span data-stu-id="0266d-113">Properties</span></span>

| <span data-ttu-id="0266d-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="0266d-114">Property</span></span> | <span data-ttu-id="0266d-115">モード</span><span class="sxs-lookup"><span data-stu-id="0266d-115">Modes</span></span> | <span data-ttu-id="0266d-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="0266d-116">Return type</span></span> | <span data-ttu-id="0266d-117">最小値</span><span class="sxs-lookup"><span data-stu-id="0266d-117">Minimum</span></span><br><span data-ttu-id="0266d-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="0266d-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="0266d-119">context</span><span class="sxs-lookup"><span data-stu-id="0266d-119">context</span></span>](office.context.md) | <span data-ttu-id="0266d-120">作成</span><span class="sxs-lookup"><span data-stu-id="0266d-120">Compose</span></span><br><span data-ttu-id="0266d-121">読み取り</span><span class="sxs-lookup"><span data-stu-id="0266d-121">Read</span></span> | [<span data-ttu-id="0266d-122">Context</span><span class="sxs-lookup"><span data-stu-id="0266d-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.5) | [<span data-ttu-id="0266d-123">1.1</span><span class="sxs-lookup"><span data-stu-id="0266d-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="0266d-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="0266d-124">Enumerations</span></span>

| <span data-ttu-id="0266d-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="0266d-125">Enumeration</span></span> | <span data-ttu-id="0266d-126">モード</span><span class="sxs-lookup"><span data-stu-id="0266d-126">Modes</span></span> | <span data-ttu-id="0266d-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="0266d-127">Return type</span></span> | <span data-ttu-id="0266d-128">最小値</span><span class="sxs-lookup"><span data-stu-id="0266d-128">Minimum</span></span><br><span data-ttu-id="0266d-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="0266d-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="0266d-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="0266d-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="0266d-131">作成</span><span class="sxs-lookup"><span data-stu-id="0266d-131">Compose</span></span><br><span data-ttu-id="0266d-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="0266d-132">Read</span></span> | <span data-ttu-id="0266d-133">文字列</span><span class="sxs-lookup"><span data-stu-id="0266d-133">String</span></span> | [<span data-ttu-id="0266d-134">1.1</span><span class="sxs-lookup"><span data-stu-id="0266d-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0266d-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="0266d-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="0266d-136">作成</span><span class="sxs-lookup"><span data-stu-id="0266d-136">Compose</span></span><br><span data-ttu-id="0266d-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="0266d-137">Read</span></span> | <span data-ttu-id="0266d-138">文字列</span><span class="sxs-lookup"><span data-stu-id="0266d-138">String</span></span> | [<span data-ttu-id="0266d-139">1.1</span><span class="sxs-lookup"><span data-stu-id="0266d-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0266d-140">EventType</span><span class="sxs-lookup"><span data-stu-id="0266d-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="0266d-141">作成</span><span class="sxs-lookup"><span data-stu-id="0266d-141">Compose</span></span><br><span data-ttu-id="0266d-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="0266d-142">Read</span></span> | <span data-ttu-id="0266d-143">文字列</span><span class="sxs-lookup"><span data-stu-id="0266d-143">String</span></span> | [<span data-ttu-id="0266d-144">1.5</span><span class="sxs-lookup"><span data-stu-id="0266d-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="0266d-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="0266d-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="0266d-146">作成</span><span class="sxs-lookup"><span data-stu-id="0266d-146">Compose</span></span><br><span data-ttu-id="0266d-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="0266d-147">Read</span></span> | <span data-ttu-id="0266d-148">文字列</span><span class="sxs-lookup"><span data-stu-id="0266d-148">String</span></span> | [<span data-ttu-id="0266d-149">1.1</span><span class="sxs-lookup"><span data-stu-id="0266d-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="0266d-150">名前空間</span><span class="sxs-lookup"><span data-stu-id="0266d-150">Namespaces</span></span>

<span data-ttu-id="0266d-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="0266d-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="0266d-152">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="0266d-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="0266d-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="0266d-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="0266d-154">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="0266d-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="0266d-155">型</span><span class="sxs-lookup"><span data-stu-id="0266d-155">Type</span></span>

*   <span data-ttu-id="0266d-156">String</span><span class="sxs-lookup"><span data-stu-id="0266d-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0266d-157">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="0266d-157">Properties:</span></span>

|<span data-ttu-id="0266d-158">名前</span><span class="sxs-lookup"><span data-stu-id="0266d-158">Name</span></span>| <span data-ttu-id="0266d-159">種類</span><span class="sxs-lookup"><span data-stu-id="0266d-159">Type</span></span>| <span data-ttu-id="0266d-160">説明</span><span class="sxs-lookup"><span data-stu-id="0266d-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="0266d-161">文字列</span><span class="sxs-lookup"><span data-stu-id="0266d-161">String</span></span>|<span data-ttu-id="0266d-162">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="0266d-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="0266d-163">文字列</span><span class="sxs-lookup"><span data-stu-id="0266d-163">String</span></span>|<span data-ttu-id="0266d-164">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="0266d-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0266d-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="0266d-165">Requirements</span></span>

|<span data-ttu-id="0266d-166">要件</span><span class="sxs-lookup"><span data-stu-id="0266d-166">Requirement</span></span>| <span data-ttu-id="0266d-167">値</span><span class="sxs-lookup"><span data-stu-id="0266d-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="0266d-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0266d-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0266d-169">1.1</span><span class="sxs-lookup"><span data-stu-id="0266d-169">1.1</span></span>|
|[<span data-ttu-id="0266d-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0266d-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0266d-171">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0266d-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="0266d-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="0266d-172">CoercionType: String</span></span>

<span data-ttu-id="0266d-173">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="0266d-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0266d-174">型</span><span class="sxs-lookup"><span data-stu-id="0266d-174">Type</span></span>

*   <span data-ttu-id="0266d-175">String</span><span class="sxs-lookup"><span data-stu-id="0266d-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0266d-176">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="0266d-176">Properties:</span></span>

|<span data-ttu-id="0266d-177">名前</span><span class="sxs-lookup"><span data-stu-id="0266d-177">Name</span></span>| <span data-ttu-id="0266d-178">種類</span><span class="sxs-lookup"><span data-stu-id="0266d-178">Type</span></span>| <span data-ttu-id="0266d-179">説明</span><span class="sxs-lookup"><span data-stu-id="0266d-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="0266d-180">文字列</span><span class="sxs-lookup"><span data-stu-id="0266d-180">String</span></span>|<span data-ttu-id="0266d-181">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="0266d-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="0266d-182">String</span><span class="sxs-lookup"><span data-stu-id="0266d-182">String</span></span>|<span data-ttu-id="0266d-183">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="0266d-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0266d-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="0266d-184">Requirements</span></span>

|<span data-ttu-id="0266d-185">要件</span><span class="sxs-lookup"><span data-stu-id="0266d-185">Requirement</span></span>| <span data-ttu-id="0266d-186">値</span><span class="sxs-lookup"><span data-stu-id="0266d-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="0266d-187">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0266d-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0266d-188">1.1</span><span class="sxs-lookup"><span data-stu-id="0266d-188">1.1</span></span>|
|[<span data-ttu-id="0266d-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0266d-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0266d-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0266d-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="0266d-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="0266d-191">EventType: String</span></span>

<span data-ttu-id="0266d-192">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="0266d-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="0266d-193">型</span><span class="sxs-lookup"><span data-stu-id="0266d-193">Type</span></span>

*   <span data-ttu-id="0266d-194">String</span><span class="sxs-lookup"><span data-stu-id="0266d-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0266d-195">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="0266d-195">Properties:</span></span>

| <span data-ttu-id="0266d-196">名前</span><span class="sxs-lookup"><span data-stu-id="0266d-196">Name</span></span> | <span data-ttu-id="0266d-197">種類</span><span class="sxs-lookup"><span data-stu-id="0266d-197">Type</span></span> | <span data-ttu-id="0266d-198">説明</span><span class="sxs-lookup"><span data-stu-id="0266d-198">Description</span></span> | <span data-ttu-id="0266d-199">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="0266d-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="0266d-200">文字列</span><span class="sxs-lookup"><span data-stu-id="0266d-200">String</span></span> | <span data-ttu-id="0266d-201">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="0266d-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="0266d-202">1.5</span><span class="sxs-lookup"><span data-stu-id="0266d-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0266d-203">Requirements</span><span class="sxs-lookup"><span data-stu-id="0266d-203">Requirements</span></span>

|<span data-ttu-id="0266d-204">要件</span><span class="sxs-lookup"><span data-stu-id="0266d-204">Requirement</span></span>| <span data-ttu-id="0266d-205">値</span><span class="sxs-lookup"><span data-stu-id="0266d-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="0266d-206">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0266d-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0266d-207">1.5</span><span class="sxs-lookup"><span data-stu-id="0266d-207">1.5</span></span> |
|[<span data-ttu-id="0266d-208">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0266d-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0266d-209">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0266d-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="0266d-210">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="0266d-210">SourceProperty: String</span></span>

<span data-ttu-id="0266d-211">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="0266d-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0266d-212">型</span><span class="sxs-lookup"><span data-stu-id="0266d-212">Type</span></span>

*   <span data-ttu-id="0266d-213">String</span><span class="sxs-lookup"><span data-stu-id="0266d-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0266d-214">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="0266d-214">Properties:</span></span>

|<span data-ttu-id="0266d-215">名前</span><span class="sxs-lookup"><span data-stu-id="0266d-215">Name</span></span>| <span data-ttu-id="0266d-216">種類</span><span class="sxs-lookup"><span data-stu-id="0266d-216">Type</span></span>| <span data-ttu-id="0266d-217">説明</span><span class="sxs-lookup"><span data-stu-id="0266d-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="0266d-218">文字列</span><span class="sxs-lookup"><span data-stu-id="0266d-218">String</span></span>|<span data-ttu-id="0266d-219">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="0266d-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="0266d-220">文字列</span><span class="sxs-lookup"><span data-stu-id="0266d-220">String</span></span>|<span data-ttu-id="0266d-221">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="0266d-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0266d-222">Requirements</span><span class="sxs-lookup"><span data-stu-id="0266d-222">Requirements</span></span>

|<span data-ttu-id="0266d-223">要件</span><span class="sxs-lookup"><span data-stu-id="0266d-223">Requirement</span></span>| <span data-ttu-id="0266d-224">値</span><span class="sxs-lookup"><span data-stu-id="0266d-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="0266d-225">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0266d-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0266d-226">1.1</span><span class="sxs-lookup"><span data-stu-id="0266d-226">1.1</span></span>|
|[<span data-ttu-id="0266d-227">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0266d-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0266d-228">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0266d-228">Compose or Read</span></span>|
