---
title: Office 名前空間-要件セット1.6
description: Outlook アドイン API の最上位レベルの名前空間のオブジェクトモデル (Mailbox API 1.6 バージョン)。
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: ae2f863e054016636ebffc3ff3925cee018036a1
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717650"
---
# <a name="office"></a><span data-ttu-id="ac66b-103">Office</span><span class="sxs-lookup"><span data-stu-id="ac66b-103">Office</span></span>

<span data-ttu-id="ac66b-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ac66b-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ac66b-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="ac66b-106">Requirements</span></span>

|<span data-ttu-id="ac66b-107">要件</span><span class="sxs-lookup"><span data-stu-id="ac66b-107">Requirement</span></span>| <span data-ttu-id="ac66b-108">値</span><span class="sxs-lookup"><span data-stu-id="ac66b-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac66b-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ac66b-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac66b-110">1.1</span><span class="sxs-lookup"><span data-stu-id="ac66b-110">1.1</span></span>|
|[<span data-ttu-id="ac66b-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ac66b-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac66b-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ac66b-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="ac66b-113">Properties</span><span class="sxs-lookup"><span data-stu-id="ac66b-113">Properties</span></span>

| <span data-ttu-id="ac66b-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="ac66b-114">Property</span></span> | <span data-ttu-id="ac66b-115">モード</span><span class="sxs-lookup"><span data-stu-id="ac66b-115">Modes</span></span> | <span data-ttu-id="ac66b-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="ac66b-116">Return type</span></span> | <span data-ttu-id="ac66b-117">最小値</span><span class="sxs-lookup"><span data-stu-id="ac66b-117">Minimum</span></span><br><span data-ttu-id="ac66b-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="ac66b-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ac66b-119">context</span><span class="sxs-lookup"><span data-stu-id="ac66b-119">context</span></span>](office.context.md) | <span data-ttu-id="ac66b-120">作成</span><span class="sxs-lookup"><span data-stu-id="ac66b-120">Compose</span></span><br><span data-ttu-id="ac66b-121">読み取り</span><span class="sxs-lookup"><span data-stu-id="ac66b-121">Read</span></span> | [<span data-ttu-id="ac66b-122">Context</span><span class="sxs-lookup"><span data-stu-id="ac66b-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6) | [<span data-ttu-id="ac66b-123">1.1</span><span class="sxs-lookup"><span data-stu-id="ac66b-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="ac66b-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="ac66b-124">Enumerations</span></span>

| <span data-ttu-id="ac66b-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="ac66b-125">Enumeration</span></span> | <span data-ttu-id="ac66b-126">モード</span><span class="sxs-lookup"><span data-stu-id="ac66b-126">Modes</span></span> | <span data-ttu-id="ac66b-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="ac66b-127">Return type</span></span> | <span data-ttu-id="ac66b-128">最小値</span><span class="sxs-lookup"><span data-stu-id="ac66b-128">Minimum</span></span><br><span data-ttu-id="ac66b-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="ac66b-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ac66b-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="ac66b-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="ac66b-131">作成</span><span class="sxs-lookup"><span data-stu-id="ac66b-131">Compose</span></span><br><span data-ttu-id="ac66b-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="ac66b-132">Read</span></span> | <span data-ttu-id="ac66b-133">文字列</span><span class="sxs-lookup"><span data-stu-id="ac66b-133">String</span></span> | [<span data-ttu-id="ac66b-134">1.1</span><span class="sxs-lookup"><span data-stu-id="ac66b-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ac66b-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="ac66b-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="ac66b-136">作成</span><span class="sxs-lookup"><span data-stu-id="ac66b-136">Compose</span></span><br><span data-ttu-id="ac66b-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="ac66b-137">Read</span></span> | <span data-ttu-id="ac66b-138">文字列</span><span class="sxs-lookup"><span data-stu-id="ac66b-138">String</span></span> | [<span data-ttu-id="ac66b-139">1.1</span><span class="sxs-lookup"><span data-stu-id="ac66b-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ac66b-140">EventType</span><span class="sxs-lookup"><span data-stu-id="ac66b-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="ac66b-141">作成</span><span class="sxs-lookup"><span data-stu-id="ac66b-141">Compose</span></span><br><span data-ttu-id="ac66b-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="ac66b-142">Read</span></span> | <span data-ttu-id="ac66b-143">文字列</span><span class="sxs-lookup"><span data-stu-id="ac66b-143">String</span></span> | [<span data-ttu-id="ac66b-144">1.5</span><span class="sxs-lookup"><span data-stu-id="ac66b-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="ac66b-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="ac66b-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="ac66b-146">作成</span><span class="sxs-lookup"><span data-stu-id="ac66b-146">Compose</span></span><br><span data-ttu-id="ac66b-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="ac66b-147">Read</span></span> | <span data-ttu-id="ac66b-148">文字列</span><span class="sxs-lookup"><span data-stu-id="ac66b-148">String</span></span> | [<span data-ttu-id="ac66b-149">1.1</span><span class="sxs-lookup"><span data-stu-id="ac66b-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="ac66b-150">名前空間</span><span class="sxs-lookup"><span data-stu-id="ac66b-150">Namespaces</span></span>

<span data-ttu-id="ac66b-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="ac66b-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="ac66b-152">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="ac66b-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="ac66b-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="ac66b-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="ac66b-154">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="ac66b-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="ac66b-155">型</span><span class="sxs-lookup"><span data-stu-id="ac66b-155">Type</span></span>

*   <span data-ttu-id="ac66b-156">String</span><span class="sxs-lookup"><span data-stu-id="ac66b-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ac66b-157">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="ac66b-157">Properties:</span></span>

|<span data-ttu-id="ac66b-158">名前</span><span class="sxs-lookup"><span data-stu-id="ac66b-158">Name</span></span>| <span data-ttu-id="ac66b-159">種類</span><span class="sxs-lookup"><span data-stu-id="ac66b-159">Type</span></span>| <span data-ttu-id="ac66b-160">説明</span><span class="sxs-lookup"><span data-stu-id="ac66b-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="ac66b-161">文字列</span><span class="sxs-lookup"><span data-stu-id="ac66b-161">String</span></span>|<span data-ttu-id="ac66b-162">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="ac66b-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="ac66b-163">文字列</span><span class="sxs-lookup"><span data-stu-id="ac66b-163">String</span></span>|<span data-ttu-id="ac66b-164">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="ac66b-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ac66b-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="ac66b-165">Requirements</span></span>

|<span data-ttu-id="ac66b-166">要件</span><span class="sxs-lookup"><span data-stu-id="ac66b-166">Requirement</span></span>| <span data-ttu-id="ac66b-167">値</span><span class="sxs-lookup"><span data-stu-id="ac66b-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac66b-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ac66b-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac66b-169">1.1</span><span class="sxs-lookup"><span data-stu-id="ac66b-169">1.1</span></span>|
|[<span data-ttu-id="ac66b-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ac66b-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac66b-171">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ac66b-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="ac66b-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="ac66b-172">CoercionType: String</span></span>

<span data-ttu-id="ac66b-173">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="ac66b-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ac66b-174">型</span><span class="sxs-lookup"><span data-stu-id="ac66b-174">Type</span></span>

*   <span data-ttu-id="ac66b-175">String</span><span class="sxs-lookup"><span data-stu-id="ac66b-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ac66b-176">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="ac66b-176">Properties:</span></span>

|<span data-ttu-id="ac66b-177">名前</span><span class="sxs-lookup"><span data-stu-id="ac66b-177">Name</span></span>| <span data-ttu-id="ac66b-178">種類</span><span class="sxs-lookup"><span data-stu-id="ac66b-178">Type</span></span>| <span data-ttu-id="ac66b-179">説明</span><span class="sxs-lookup"><span data-stu-id="ac66b-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="ac66b-180">文字列</span><span class="sxs-lookup"><span data-stu-id="ac66b-180">String</span></span>|<span data-ttu-id="ac66b-181">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="ac66b-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="ac66b-182">String</span><span class="sxs-lookup"><span data-stu-id="ac66b-182">String</span></span>|<span data-ttu-id="ac66b-183">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="ac66b-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ac66b-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="ac66b-184">Requirements</span></span>

|<span data-ttu-id="ac66b-185">要件</span><span class="sxs-lookup"><span data-stu-id="ac66b-185">Requirement</span></span>| <span data-ttu-id="ac66b-186">値</span><span class="sxs-lookup"><span data-stu-id="ac66b-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac66b-187">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ac66b-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac66b-188">1.1</span><span class="sxs-lookup"><span data-stu-id="ac66b-188">1.1</span></span>|
|[<span data-ttu-id="ac66b-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ac66b-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac66b-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ac66b-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="ac66b-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="ac66b-191">EventType: String</span></span>

<span data-ttu-id="ac66b-192">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="ac66b-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="ac66b-193">型</span><span class="sxs-lookup"><span data-stu-id="ac66b-193">Type</span></span>

*   <span data-ttu-id="ac66b-194">String</span><span class="sxs-lookup"><span data-stu-id="ac66b-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ac66b-195">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="ac66b-195">Properties:</span></span>

| <span data-ttu-id="ac66b-196">名前</span><span class="sxs-lookup"><span data-stu-id="ac66b-196">Name</span></span> | <span data-ttu-id="ac66b-197">種類</span><span class="sxs-lookup"><span data-stu-id="ac66b-197">Type</span></span> | <span data-ttu-id="ac66b-198">説明</span><span class="sxs-lookup"><span data-stu-id="ac66b-198">Description</span></span> | <span data-ttu-id="ac66b-199">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="ac66b-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="ac66b-200">文字列</span><span class="sxs-lookup"><span data-stu-id="ac66b-200">String</span></span> | <span data-ttu-id="ac66b-201">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="ac66b-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="ac66b-202">1.5</span><span class="sxs-lookup"><span data-stu-id="ac66b-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ac66b-203">Requirements</span><span class="sxs-lookup"><span data-stu-id="ac66b-203">Requirements</span></span>

|<span data-ttu-id="ac66b-204">要件</span><span class="sxs-lookup"><span data-stu-id="ac66b-204">Requirement</span></span>| <span data-ttu-id="ac66b-205">値</span><span class="sxs-lookup"><span data-stu-id="ac66b-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac66b-206">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ac66b-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac66b-207">1.5</span><span class="sxs-lookup"><span data-stu-id="ac66b-207">1.5</span></span> |
|[<span data-ttu-id="ac66b-208">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ac66b-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac66b-209">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ac66b-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="ac66b-210">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="ac66b-210">SourceProperty: String</span></span>

<span data-ttu-id="ac66b-211">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="ac66b-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ac66b-212">型</span><span class="sxs-lookup"><span data-stu-id="ac66b-212">Type</span></span>

*   <span data-ttu-id="ac66b-213">String</span><span class="sxs-lookup"><span data-stu-id="ac66b-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ac66b-214">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="ac66b-214">Properties:</span></span>

|<span data-ttu-id="ac66b-215">名前</span><span class="sxs-lookup"><span data-stu-id="ac66b-215">Name</span></span>| <span data-ttu-id="ac66b-216">種類</span><span class="sxs-lookup"><span data-stu-id="ac66b-216">Type</span></span>| <span data-ttu-id="ac66b-217">説明</span><span class="sxs-lookup"><span data-stu-id="ac66b-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="ac66b-218">文字列</span><span class="sxs-lookup"><span data-stu-id="ac66b-218">String</span></span>|<span data-ttu-id="ac66b-219">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="ac66b-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="ac66b-220">文字列</span><span class="sxs-lookup"><span data-stu-id="ac66b-220">String</span></span>|<span data-ttu-id="ac66b-221">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="ac66b-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ac66b-222">Requirements</span><span class="sxs-lookup"><span data-stu-id="ac66b-222">Requirements</span></span>

|<span data-ttu-id="ac66b-223">要件</span><span class="sxs-lookup"><span data-stu-id="ac66b-223">Requirement</span></span>| <span data-ttu-id="ac66b-224">値</span><span class="sxs-lookup"><span data-stu-id="ac66b-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac66b-225">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ac66b-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac66b-226">1.1</span><span class="sxs-lookup"><span data-stu-id="ac66b-226">1.1</span></span>|
|[<span data-ttu-id="ac66b-227">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ac66b-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac66b-228">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ac66b-228">Compose or Read</span></span>|
