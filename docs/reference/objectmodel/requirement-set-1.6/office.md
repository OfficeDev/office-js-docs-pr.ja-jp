---
title: Office 名前空間-要件セット1.6
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 0a6360ff7f4e397b878d9a3f744bdbe58347c558
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163662"
---
# <a name="office"></a><span data-ttu-id="a2d0d-102">Office</span><span class="sxs-lookup"><span data-stu-id="a2d0d-102">Office</span></span>

<span data-ttu-id="a2d0d-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a2d0d-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a2d0d-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="a2d0d-105">Requirements</span></span>

|<span data-ttu-id="a2d0d-106">要件</span><span class="sxs-lookup"><span data-stu-id="a2d0d-106">Requirement</span></span>| <span data-ttu-id="a2d0d-107">値</span><span class="sxs-lookup"><span data-stu-id="a2d0d-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a2d0d-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a2d0d-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a2d0d-109">1.1</span><span class="sxs-lookup"><span data-stu-id="a2d0d-109">1.1</span></span>|
|[<span data-ttu-id="a2d0d-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a2d0d-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a2d0d-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="a2d0d-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="a2d0d-112">Properties</span><span class="sxs-lookup"><span data-stu-id="a2d0d-112">Properties</span></span>

| <span data-ttu-id="a2d0d-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="a2d0d-113">Property</span></span> | <span data-ttu-id="a2d0d-114">モード</span><span class="sxs-lookup"><span data-stu-id="a2d0d-114">Modes</span></span> | <span data-ttu-id="a2d0d-115">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="a2d0d-115">Return type</span></span> | <span data-ttu-id="a2d0d-116">最小値</span><span class="sxs-lookup"><span data-stu-id="a2d0d-116">Minimum</span></span><br><span data-ttu-id="a2d0d-117">要件セット</span><span class="sxs-lookup"><span data-stu-id="a2d0d-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="a2d0d-118">context</span><span class="sxs-lookup"><span data-stu-id="a2d0d-118">context</span></span>](office.context.md) | <span data-ttu-id="a2d0d-119">作成</span><span class="sxs-lookup"><span data-stu-id="a2d0d-119">Compose</span></span><br><span data-ttu-id="a2d0d-120">読み取り</span><span class="sxs-lookup"><span data-stu-id="a2d0d-120">Read</span></span> | [<span data-ttu-id="a2d0d-121">Context</span><span class="sxs-lookup"><span data-stu-id="a2d0d-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6) | [<span data-ttu-id="a2d0d-122">1.1</span><span class="sxs-lookup"><span data-stu-id="a2d0d-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="a2d0d-123">列挙型</span><span class="sxs-lookup"><span data-stu-id="a2d0d-123">Enumerations</span></span>

| <span data-ttu-id="a2d0d-124">列挙体</span><span class="sxs-lookup"><span data-stu-id="a2d0d-124">Enumeration</span></span> | <span data-ttu-id="a2d0d-125">モード</span><span class="sxs-lookup"><span data-stu-id="a2d0d-125">Modes</span></span> | <span data-ttu-id="a2d0d-126">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="a2d0d-126">Return type</span></span> | <span data-ttu-id="a2d0d-127">最小値</span><span class="sxs-lookup"><span data-stu-id="a2d0d-127">Minimum</span></span><br><span data-ttu-id="a2d0d-128">要件セット</span><span class="sxs-lookup"><span data-stu-id="a2d0d-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="a2d0d-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="a2d0d-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="a2d0d-130">作成</span><span class="sxs-lookup"><span data-stu-id="a2d0d-130">Compose</span></span><br><span data-ttu-id="a2d0d-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="a2d0d-131">Read</span></span> | <span data-ttu-id="a2d0d-132">文字列</span><span class="sxs-lookup"><span data-stu-id="a2d0d-132">String</span></span> | [<span data-ttu-id="a2d0d-133">1.1</span><span class="sxs-lookup"><span data-stu-id="a2d0d-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a2d0d-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="a2d0d-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="a2d0d-135">作成</span><span class="sxs-lookup"><span data-stu-id="a2d0d-135">Compose</span></span><br><span data-ttu-id="a2d0d-136">読み取り</span><span class="sxs-lookup"><span data-stu-id="a2d0d-136">Read</span></span> | <span data-ttu-id="a2d0d-137">文字列</span><span class="sxs-lookup"><span data-stu-id="a2d0d-137">String</span></span> | [<span data-ttu-id="a2d0d-138">1.1</span><span class="sxs-lookup"><span data-stu-id="a2d0d-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a2d0d-139">EventType</span><span class="sxs-lookup"><span data-stu-id="a2d0d-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="a2d0d-140">作成</span><span class="sxs-lookup"><span data-stu-id="a2d0d-140">Compose</span></span><br><span data-ttu-id="a2d0d-141">読み取り</span><span class="sxs-lookup"><span data-stu-id="a2d0d-141">Read</span></span> | <span data-ttu-id="a2d0d-142">文字列</span><span class="sxs-lookup"><span data-stu-id="a2d0d-142">String</span></span> | [<span data-ttu-id="a2d0d-143">1.5</span><span class="sxs-lookup"><span data-stu-id="a2d0d-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="a2d0d-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="a2d0d-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="a2d0d-145">作成</span><span class="sxs-lookup"><span data-stu-id="a2d0d-145">Compose</span></span><br><span data-ttu-id="a2d0d-146">読み取り</span><span class="sxs-lookup"><span data-stu-id="a2d0d-146">Read</span></span> | <span data-ttu-id="a2d0d-147">文字列</span><span class="sxs-lookup"><span data-stu-id="a2d0d-147">String</span></span> | [<span data-ttu-id="a2d0d-148">1.1</span><span class="sxs-lookup"><span data-stu-id="a2d0d-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="a2d0d-149">名前空間</span><span class="sxs-lookup"><span data-stu-id="a2d0d-149">Namespaces</span></span>

<span data-ttu-id="a2d0d-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): `ItemType`、、、、、 `EntityType` `AttachmentType` `RecipientType` `ResponseType`など、多数の Outlook 固有の列挙を含み`ItemNotificationMessageType`ます。</span><span class="sxs-lookup"><span data-stu-id="a2d0d-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="a2d0d-151">列挙の詳細</span><span class="sxs-lookup"><span data-stu-id="a2d0d-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="a2d0d-152">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="a2d0d-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="a2d0d-153">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="a2d0d-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="a2d0d-154">型</span><span class="sxs-lookup"><span data-stu-id="a2d0d-154">Type</span></span>

*   <span data-ttu-id="a2d0d-155">String</span><span class="sxs-lookup"><span data-stu-id="a2d0d-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a2d0d-156">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="a2d0d-156">Properties:</span></span>

|<span data-ttu-id="a2d0d-157">名前</span><span class="sxs-lookup"><span data-stu-id="a2d0d-157">Name</span></span>| <span data-ttu-id="a2d0d-158">種類</span><span class="sxs-lookup"><span data-stu-id="a2d0d-158">Type</span></span>| <span data-ttu-id="a2d0d-159">説明</span><span class="sxs-lookup"><span data-stu-id="a2d0d-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="a2d0d-160">文字列</span><span class="sxs-lookup"><span data-stu-id="a2d0d-160">String</span></span>|<span data-ttu-id="a2d0d-161">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="a2d0d-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="a2d0d-162">文字列</span><span class="sxs-lookup"><span data-stu-id="a2d0d-162">String</span></span>|<span data-ttu-id="a2d0d-163">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="a2d0d-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a2d0d-164">Requirements</span><span class="sxs-lookup"><span data-stu-id="a2d0d-164">Requirements</span></span>

|<span data-ttu-id="a2d0d-165">要件</span><span class="sxs-lookup"><span data-stu-id="a2d0d-165">Requirement</span></span>| <span data-ttu-id="a2d0d-166">値</span><span class="sxs-lookup"><span data-stu-id="a2d0d-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="a2d0d-167">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a2d0d-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a2d0d-168">1.1</span><span class="sxs-lookup"><span data-stu-id="a2d0d-168">1.1</span></span>|
|[<span data-ttu-id="a2d0d-169">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a2d0d-169">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a2d0d-170">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="a2d0d-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="a2d0d-171">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="a2d0d-171">CoercionType: String</span></span>

<span data-ttu-id="a2d0d-172">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="a2d0d-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a2d0d-173">型</span><span class="sxs-lookup"><span data-stu-id="a2d0d-173">Type</span></span>

*   <span data-ttu-id="a2d0d-174">String</span><span class="sxs-lookup"><span data-stu-id="a2d0d-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a2d0d-175">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="a2d0d-175">Properties:</span></span>

|<span data-ttu-id="a2d0d-176">名前</span><span class="sxs-lookup"><span data-stu-id="a2d0d-176">Name</span></span>| <span data-ttu-id="a2d0d-177">種類</span><span class="sxs-lookup"><span data-stu-id="a2d0d-177">Type</span></span>| <span data-ttu-id="a2d0d-178">説明</span><span class="sxs-lookup"><span data-stu-id="a2d0d-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="a2d0d-179">文字列</span><span class="sxs-lookup"><span data-stu-id="a2d0d-179">String</span></span>|<span data-ttu-id="a2d0d-180">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="a2d0d-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="a2d0d-181">String</span><span class="sxs-lookup"><span data-stu-id="a2d0d-181">String</span></span>|<span data-ttu-id="a2d0d-182">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="a2d0d-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a2d0d-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="a2d0d-183">Requirements</span></span>

|<span data-ttu-id="a2d0d-184">要件</span><span class="sxs-lookup"><span data-stu-id="a2d0d-184">Requirement</span></span>| <span data-ttu-id="a2d0d-185">値</span><span class="sxs-lookup"><span data-stu-id="a2d0d-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="a2d0d-186">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a2d0d-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a2d0d-187">1.1</span><span class="sxs-lookup"><span data-stu-id="a2d0d-187">1.1</span></span>|
|[<span data-ttu-id="a2d0d-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a2d0d-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a2d0d-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="a2d0d-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="a2d0d-190">EventType: String</span><span class="sxs-lookup"><span data-stu-id="a2d0d-190">EventType: String</span></span>

<span data-ttu-id="a2d0d-191">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="a2d0d-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="a2d0d-192">型</span><span class="sxs-lookup"><span data-stu-id="a2d0d-192">Type</span></span>

*   <span data-ttu-id="a2d0d-193">String</span><span class="sxs-lookup"><span data-stu-id="a2d0d-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a2d0d-194">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="a2d0d-194">Properties:</span></span>

| <span data-ttu-id="a2d0d-195">名前</span><span class="sxs-lookup"><span data-stu-id="a2d0d-195">Name</span></span> | <span data-ttu-id="a2d0d-196">種類</span><span class="sxs-lookup"><span data-stu-id="a2d0d-196">Type</span></span> | <span data-ttu-id="a2d0d-197">説明</span><span class="sxs-lookup"><span data-stu-id="a2d0d-197">Description</span></span> | <span data-ttu-id="a2d0d-198">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="a2d0d-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="a2d0d-199">文字列</span><span class="sxs-lookup"><span data-stu-id="a2d0d-199">String</span></span> | <span data-ttu-id="a2d0d-200">作業ウィンドウが固定されている間、別の Outlook アイテムが選択され、表示することができます。</span><span class="sxs-lookup"><span data-stu-id="a2d0d-200">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="a2d0d-201">1.5</span><span class="sxs-lookup"><span data-stu-id="a2d0d-201">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a2d0d-202">Requirements</span><span class="sxs-lookup"><span data-stu-id="a2d0d-202">Requirements</span></span>

|<span data-ttu-id="a2d0d-203">要件</span><span class="sxs-lookup"><span data-stu-id="a2d0d-203">Requirement</span></span>| <span data-ttu-id="a2d0d-204">値</span><span class="sxs-lookup"><span data-stu-id="a2d0d-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="a2d0d-205">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a2d0d-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a2d0d-206">1.5</span><span class="sxs-lookup"><span data-stu-id="a2d0d-206">1.5</span></span> |
|[<span data-ttu-id="a2d0d-207">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a2d0d-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a2d0d-208">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="a2d0d-208">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="a2d0d-209">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="a2d0d-209">SourceProperty: String</span></span>

<span data-ttu-id="a2d0d-210">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="a2d0d-210">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a2d0d-211">型</span><span class="sxs-lookup"><span data-stu-id="a2d0d-211">Type</span></span>

*   <span data-ttu-id="a2d0d-212">String</span><span class="sxs-lookup"><span data-stu-id="a2d0d-212">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a2d0d-213">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="a2d0d-213">Properties:</span></span>

|<span data-ttu-id="a2d0d-214">名前</span><span class="sxs-lookup"><span data-stu-id="a2d0d-214">Name</span></span>| <span data-ttu-id="a2d0d-215">種類</span><span class="sxs-lookup"><span data-stu-id="a2d0d-215">Type</span></span>| <span data-ttu-id="a2d0d-216">説明</span><span class="sxs-lookup"><span data-stu-id="a2d0d-216">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="a2d0d-217">文字列</span><span class="sxs-lookup"><span data-stu-id="a2d0d-217">String</span></span>|<span data-ttu-id="a2d0d-218">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="a2d0d-218">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="a2d0d-219">文字列</span><span class="sxs-lookup"><span data-stu-id="a2d0d-219">String</span></span>|<span data-ttu-id="a2d0d-220">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="a2d0d-220">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a2d0d-221">Requirements</span><span class="sxs-lookup"><span data-stu-id="a2d0d-221">Requirements</span></span>

|<span data-ttu-id="a2d0d-222">要件</span><span class="sxs-lookup"><span data-stu-id="a2d0d-222">Requirement</span></span>| <span data-ttu-id="a2d0d-223">値</span><span class="sxs-lookup"><span data-stu-id="a2d0d-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="a2d0d-224">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a2d0d-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a2d0d-225">1.1</span><span class="sxs-lookup"><span data-stu-id="a2d0d-225">1.1</span></span>|
|[<span data-ttu-id="a2d0d-226">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a2d0d-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a2d0d-227">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="a2d0d-227">Compose or Read</span></span>|
