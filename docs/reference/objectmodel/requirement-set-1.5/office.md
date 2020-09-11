---
title: Office 名前空間-要件セット1.5
description: メールボックス API 要件セット1.5 を使用した Outlook アドインで使用可能な Office 名前空間メンバー。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 141fd124ba5778a5ae576c7b4cd2c749a9c4bd6f
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430598"
---
# <a name="office-mailbox-requirement-set-15"></a><span data-ttu-id="2d2a0-103">Office (メールボックス要件セット 1.5)</span><span class="sxs-lookup"><span data-stu-id="2d2a0-103">Office (Mailbox requirement set 1.5)</span></span>

<span data-ttu-id="2d2a0-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2d2a0-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="2d2a0-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="2d2a0-106">Requirements</span></span>

|<span data-ttu-id="2d2a0-107">要件</span><span class="sxs-lookup"><span data-stu-id="2d2a0-107">Requirement</span></span>| <span data-ttu-id="2d2a0-108">値</span><span class="sxs-lookup"><span data-stu-id="2d2a0-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d2a0-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2d2a0-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2d2a0-110">1.1</span><span class="sxs-lookup"><span data-stu-id="2d2a0-110">1.1</span></span>|
|[<span data-ttu-id="2d2a0-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2d2a0-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2d2a0-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2d2a0-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="2d2a0-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="2d2a0-113">Properties</span></span>

| <span data-ttu-id="2d2a0-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="2d2a0-114">Property</span></span> | <span data-ttu-id="2d2a0-115">モード</span><span class="sxs-lookup"><span data-stu-id="2d2a0-115">Modes</span></span> | <span data-ttu-id="2d2a0-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="2d2a0-116">Return type</span></span> | <span data-ttu-id="2d2a0-117">最小値</span><span class="sxs-lookup"><span data-stu-id="2d2a0-117">Minimum</span></span><br><span data-ttu-id="2d2a0-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="2d2a0-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2d2a0-119">context</span><span class="sxs-lookup"><span data-stu-id="2d2a0-119">context</span></span>](office.context.md) | <span data-ttu-id="2d2a0-120">作成</span><span class="sxs-lookup"><span data-stu-id="2d2a0-120">Compose</span></span><br><span data-ttu-id="2d2a0-121">読み取り</span><span class="sxs-lookup"><span data-stu-id="2d2a0-121">Read</span></span> | [<span data-ttu-id="2d2a0-122">Context</span><span class="sxs-lookup"><span data-stu-id="2d2a0-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="2d2a0-123">1.1</span><span class="sxs-lookup"><span data-stu-id="2d2a0-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="2d2a0-124">列挙型</span><span class="sxs-lookup"><span data-stu-id="2d2a0-124">Enumerations</span></span>

| <span data-ttu-id="2d2a0-125">列挙体</span><span class="sxs-lookup"><span data-stu-id="2d2a0-125">Enumeration</span></span> | <span data-ttu-id="2d2a0-126">モード</span><span class="sxs-lookup"><span data-stu-id="2d2a0-126">Modes</span></span> | <span data-ttu-id="2d2a0-127">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="2d2a0-127">Return type</span></span> | <span data-ttu-id="2d2a0-128">最小値</span><span class="sxs-lookup"><span data-stu-id="2d2a0-128">Minimum</span></span><br><span data-ttu-id="2d2a0-129">要件セット</span><span class="sxs-lookup"><span data-stu-id="2d2a0-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2d2a0-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="2d2a0-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="2d2a0-131">作成</span><span class="sxs-lookup"><span data-stu-id="2d2a0-131">Compose</span></span><br><span data-ttu-id="2d2a0-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="2d2a0-132">Read</span></span> | <span data-ttu-id="2d2a0-133">文字列</span><span class="sxs-lookup"><span data-stu-id="2d2a0-133">String</span></span> | [<span data-ttu-id="2d2a0-134">1.1</span><span class="sxs-lookup"><span data-stu-id="2d2a0-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2d2a0-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="2d2a0-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="2d2a0-136">作成</span><span class="sxs-lookup"><span data-stu-id="2d2a0-136">Compose</span></span><br><span data-ttu-id="2d2a0-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="2d2a0-137">Read</span></span> | <span data-ttu-id="2d2a0-138">文字列</span><span class="sxs-lookup"><span data-stu-id="2d2a0-138">String</span></span> | [<span data-ttu-id="2d2a0-139">1.1</span><span class="sxs-lookup"><span data-stu-id="2d2a0-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2d2a0-140">EventType</span><span class="sxs-lookup"><span data-stu-id="2d2a0-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="2d2a0-141">作成</span><span class="sxs-lookup"><span data-stu-id="2d2a0-141">Compose</span></span><br><span data-ttu-id="2d2a0-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="2d2a0-142">Read</span></span> | <span data-ttu-id="2d2a0-143">文字列</span><span class="sxs-lookup"><span data-stu-id="2d2a0-143">String</span></span> | [<span data-ttu-id="2d2a0-144">1.5</span><span class="sxs-lookup"><span data-stu-id="2d2a0-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="2d2a0-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="2d2a0-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="2d2a0-146">作成</span><span class="sxs-lookup"><span data-stu-id="2d2a0-146">Compose</span></span><br><span data-ttu-id="2d2a0-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="2d2a0-147">Read</span></span> | <span data-ttu-id="2d2a0-148">文字列</span><span class="sxs-lookup"><span data-stu-id="2d2a0-148">String</span></span> | [<span data-ttu-id="2d2a0-149">1.1</span><span class="sxs-lookup"><span data-stu-id="2d2a0-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="2d2a0-150">名前空間</span><span class="sxs-lookup"><span data-stu-id="2d2a0-150">Namespaces</span></span>

<span data-ttu-id="2d2a0-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5&preserve-view=true):、、、、、など、多数の Outlook 固有の列挙を含み `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` ます。</span><span class="sxs-lookup"><span data-stu-id="2d2a0-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="2d2a0-152">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="2d2a0-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="2d2a0-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="2d2a0-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="2d2a0-154">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="2d2a0-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="2d2a0-155">型</span><span class="sxs-lookup"><span data-stu-id="2d2a0-155">Type</span></span>

*   <span data-ttu-id="2d2a0-156">String</span><span class="sxs-lookup"><span data-stu-id="2d2a0-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2d2a0-157">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="2d2a0-157">Properties:</span></span>

|<span data-ttu-id="2d2a0-158">名前</span><span class="sxs-lookup"><span data-stu-id="2d2a0-158">Name</span></span>| <span data-ttu-id="2d2a0-159">種類</span><span class="sxs-lookup"><span data-stu-id="2d2a0-159">Type</span></span>| <span data-ttu-id="2d2a0-160">説明</span><span class="sxs-lookup"><span data-stu-id="2d2a0-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="2d2a0-161">文字列</span><span class="sxs-lookup"><span data-stu-id="2d2a0-161">String</span></span>|<span data-ttu-id="2d2a0-162">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="2d2a0-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="2d2a0-163">String</span><span class="sxs-lookup"><span data-stu-id="2d2a0-163">String</span></span>|<span data-ttu-id="2d2a0-164">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="2d2a0-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2d2a0-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="2d2a0-165">Requirements</span></span>

|<span data-ttu-id="2d2a0-166">要件</span><span class="sxs-lookup"><span data-stu-id="2d2a0-166">Requirement</span></span>| <span data-ttu-id="2d2a0-167">値</span><span class="sxs-lookup"><span data-stu-id="2d2a0-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d2a0-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2d2a0-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2d2a0-169">1.1</span><span class="sxs-lookup"><span data-stu-id="2d2a0-169">1.1</span></span>|
|[<span data-ttu-id="2d2a0-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2d2a0-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2d2a0-171">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2d2a0-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="2d2a0-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="2d2a0-172">CoercionType: String</span></span>

<span data-ttu-id="2d2a0-173">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="2d2a0-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2d2a0-174">型</span><span class="sxs-lookup"><span data-stu-id="2d2a0-174">Type</span></span>

*   <span data-ttu-id="2d2a0-175">String</span><span class="sxs-lookup"><span data-stu-id="2d2a0-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2d2a0-176">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="2d2a0-176">Properties:</span></span>

|<span data-ttu-id="2d2a0-177">名前</span><span class="sxs-lookup"><span data-stu-id="2d2a0-177">Name</span></span>| <span data-ttu-id="2d2a0-178">種類</span><span class="sxs-lookup"><span data-stu-id="2d2a0-178">Type</span></span>| <span data-ttu-id="2d2a0-179">説明</span><span class="sxs-lookup"><span data-stu-id="2d2a0-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="2d2a0-180">文字列</span><span class="sxs-lookup"><span data-stu-id="2d2a0-180">String</span></span>|<span data-ttu-id="2d2a0-181">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="2d2a0-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="2d2a0-182">String</span><span class="sxs-lookup"><span data-stu-id="2d2a0-182">String</span></span>|<span data-ttu-id="2d2a0-183">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="2d2a0-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2d2a0-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="2d2a0-184">Requirements</span></span>

|<span data-ttu-id="2d2a0-185">要件</span><span class="sxs-lookup"><span data-stu-id="2d2a0-185">Requirement</span></span>| <span data-ttu-id="2d2a0-186">値</span><span class="sxs-lookup"><span data-stu-id="2d2a0-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d2a0-187">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2d2a0-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2d2a0-188">1.1</span><span class="sxs-lookup"><span data-stu-id="2d2a0-188">1.1</span></span>|
|[<span data-ttu-id="2d2a0-189">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2d2a0-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2d2a0-190">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2d2a0-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="2d2a0-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="2d2a0-191">EventType: String</span></span>

<span data-ttu-id="2d2a0-192">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="2d2a0-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="2d2a0-193">型</span><span class="sxs-lookup"><span data-stu-id="2d2a0-193">Type</span></span>

*   <span data-ttu-id="2d2a0-194">String</span><span class="sxs-lookup"><span data-stu-id="2d2a0-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2d2a0-195">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="2d2a0-195">Properties:</span></span>

| <span data-ttu-id="2d2a0-196">名前</span><span class="sxs-lookup"><span data-stu-id="2d2a0-196">Name</span></span> | <span data-ttu-id="2d2a0-197">種類</span><span class="sxs-lookup"><span data-stu-id="2d2a0-197">Type</span></span> | <span data-ttu-id="2d2a0-198">説明</span><span class="sxs-lookup"><span data-stu-id="2d2a0-198">Description</span></span> | <span data-ttu-id="2d2a0-199">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="2d2a0-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="2d2a0-200">文字列</span><span class="sxs-lookup"><span data-stu-id="2d2a0-200">String</span></span> | <span data-ttu-id="2d2a0-201">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="2d2a0-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="2d2a0-202">1.5</span><span class="sxs-lookup"><span data-stu-id="2d2a0-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2d2a0-203">Requirements</span><span class="sxs-lookup"><span data-stu-id="2d2a0-203">Requirements</span></span>

|<span data-ttu-id="2d2a0-204">要件</span><span class="sxs-lookup"><span data-stu-id="2d2a0-204">Requirement</span></span>| <span data-ttu-id="2d2a0-205">値</span><span class="sxs-lookup"><span data-stu-id="2d2a0-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d2a0-206">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2d2a0-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2d2a0-207">1.5</span><span class="sxs-lookup"><span data-stu-id="2d2a0-207">1.5</span></span> |
|[<span data-ttu-id="2d2a0-208">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2d2a0-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2d2a0-209">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2d2a0-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="2d2a0-210">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="2d2a0-210">SourceProperty: String</span></span>

<span data-ttu-id="2d2a0-211">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="2d2a0-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2d2a0-212">型</span><span class="sxs-lookup"><span data-stu-id="2d2a0-212">Type</span></span>

*   <span data-ttu-id="2d2a0-213">String</span><span class="sxs-lookup"><span data-stu-id="2d2a0-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2d2a0-214">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="2d2a0-214">Properties:</span></span>

|<span data-ttu-id="2d2a0-215">名前</span><span class="sxs-lookup"><span data-stu-id="2d2a0-215">Name</span></span>| <span data-ttu-id="2d2a0-216">種類</span><span class="sxs-lookup"><span data-stu-id="2d2a0-216">Type</span></span>| <span data-ttu-id="2d2a0-217">説明</span><span class="sxs-lookup"><span data-stu-id="2d2a0-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="2d2a0-218">文字列</span><span class="sxs-lookup"><span data-stu-id="2d2a0-218">String</span></span>|<span data-ttu-id="2d2a0-219">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="2d2a0-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="2d2a0-220">String</span><span class="sxs-lookup"><span data-stu-id="2d2a0-220">String</span></span>|<span data-ttu-id="2d2a0-221">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="2d2a0-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2d2a0-222">Requirements</span><span class="sxs-lookup"><span data-stu-id="2d2a0-222">Requirements</span></span>

|<span data-ttu-id="2d2a0-223">要件</span><span class="sxs-lookup"><span data-stu-id="2d2a0-223">Requirement</span></span>| <span data-ttu-id="2d2a0-224">値</span><span class="sxs-lookup"><span data-stu-id="2d2a0-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d2a0-225">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2d2a0-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2d2a0-226">1.1</span><span class="sxs-lookup"><span data-stu-id="2d2a0-226">1.1</span></span>|
|[<span data-ttu-id="2d2a0-227">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2d2a0-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2d2a0-228">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2d2a0-228">Compose or Read</span></span>|
