---
title: Office 名前空間-要件セット1.3
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: ef01b7da3d447af852a5558853e0902eab815dd3
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871222"
---
# <a name="office"></a><span data-ttu-id="375c8-102">Office</span><span class="sxs-lookup"><span data-stu-id="375c8-102">Office</span></span>

<span data-ttu-id="375c8-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="375c8-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="375c8-105">要件</span><span class="sxs-lookup"><span data-stu-id="375c8-105">Requirements</span></span>

|<span data-ttu-id="375c8-106">要件</span><span class="sxs-lookup"><span data-stu-id="375c8-106">Requirement</span></span>| <span data-ttu-id="375c8-107">値</span><span class="sxs-lookup"><span data-stu-id="375c8-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="375c8-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="375c8-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="375c8-109">1.0</span><span class="sxs-lookup"><span data-stu-id="375c8-109">1.0</span></span>|
|[<span data-ttu-id="375c8-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="375c8-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="375c8-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="375c8-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="375c8-112">名前空間</span><span class="sxs-lookup"><span data-stu-id="375c8-112">Namespaces</span></span>

<span data-ttu-id="375c8-113">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="375c8-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="375c8-114">[MailboxEnums](/javascript/api/outlook_1_3/office.mailboxenums.attachmenttype): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="375c8-114">[MailboxEnums](/javascript/api/outlook_1_3/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="375c8-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="375c8-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="375c8-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="375c8-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="375c8-117">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="375c8-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="375c8-118">型</span><span class="sxs-lookup"><span data-stu-id="375c8-118">Type</span></span>

*   <span data-ttu-id="375c8-119">String</span><span class="sxs-lookup"><span data-stu-id="375c8-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="375c8-120">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="375c8-120">Properties:</span></span>

|<span data-ttu-id="375c8-121">名前</span><span class="sxs-lookup"><span data-stu-id="375c8-121">Name</span></span>| <span data-ttu-id="375c8-122">種類</span><span class="sxs-lookup"><span data-stu-id="375c8-122">Type</span></span>| <span data-ttu-id="375c8-123">説明</span><span class="sxs-lookup"><span data-stu-id="375c8-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="375c8-124">String</span><span class="sxs-lookup"><span data-stu-id="375c8-124">String</span></span>|<span data-ttu-id="375c8-125">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="375c8-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="375c8-126">String</span><span class="sxs-lookup"><span data-stu-id="375c8-126">String</span></span>|<span data-ttu-id="375c8-127">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="375c8-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="375c8-128">要件</span><span class="sxs-lookup"><span data-stu-id="375c8-128">Requirements</span></span>

|<span data-ttu-id="375c8-129">要件</span><span class="sxs-lookup"><span data-stu-id="375c8-129">Requirement</span></span>| <span data-ttu-id="375c8-130">値</span><span class="sxs-lookup"><span data-stu-id="375c8-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="375c8-131">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="375c8-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="375c8-132">1.0</span><span class="sxs-lookup"><span data-stu-id="375c8-132">1.0</span></span>|
|[<span data-ttu-id="375c8-133">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="375c8-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="375c8-134">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="375c8-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="375c8-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="375c8-135">CoercionType :String</span></span>

<span data-ttu-id="375c8-136">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="375c8-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="375c8-137">型</span><span class="sxs-lookup"><span data-stu-id="375c8-137">Type</span></span>

*   <span data-ttu-id="375c8-138">String</span><span class="sxs-lookup"><span data-stu-id="375c8-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="375c8-139">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="375c8-139">Properties:</span></span>

|<span data-ttu-id="375c8-140">名前</span><span class="sxs-lookup"><span data-stu-id="375c8-140">Name</span></span>| <span data-ttu-id="375c8-141">種類</span><span class="sxs-lookup"><span data-stu-id="375c8-141">Type</span></span>| <span data-ttu-id="375c8-142">説明</span><span class="sxs-lookup"><span data-stu-id="375c8-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="375c8-143">String</span><span class="sxs-lookup"><span data-stu-id="375c8-143">String</span></span>|<span data-ttu-id="375c8-144">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="375c8-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="375c8-145">String</span><span class="sxs-lookup"><span data-stu-id="375c8-145">String</span></span>|<span data-ttu-id="375c8-146">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="375c8-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="375c8-147">要件</span><span class="sxs-lookup"><span data-stu-id="375c8-147">Requirements</span></span>

|<span data-ttu-id="375c8-148">要件</span><span class="sxs-lookup"><span data-stu-id="375c8-148">Requirement</span></span>| <span data-ttu-id="375c8-149">値</span><span class="sxs-lookup"><span data-stu-id="375c8-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="375c8-150">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="375c8-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="375c8-151">1.0</span><span class="sxs-lookup"><span data-stu-id="375c8-151">1.0</span></span>|
|[<span data-ttu-id="375c8-152">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="375c8-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="375c8-153">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="375c8-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="375c8-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="375c8-154">SourceProperty :String</span></span>

<span data-ttu-id="375c8-155">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="375c8-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="375c8-156">型</span><span class="sxs-lookup"><span data-stu-id="375c8-156">Type</span></span>

*   <span data-ttu-id="375c8-157">String</span><span class="sxs-lookup"><span data-stu-id="375c8-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="375c8-158">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="375c8-158">Properties:</span></span>

|<span data-ttu-id="375c8-159">名前</span><span class="sxs-lookup"><span data-stu-id="375c8-159">Name</span></span>| <span data-ttu-id="375c8-160">種類</span><span class="sxs-lookup"><span data-stu-id="375c8-160">Type</span></span>| <span data-ttu-id="375c8-161">説明</span><span class="sxs-lookup"><span data-stu-id="375c8-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="375c8-162">String</span><span class="sxs-lookup"><span data-stu-id="375c8-162">String</span></span>|<span data-ttu-id="375c8-163">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="375c8-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="375c8-164">String</span><span class="sxs-lookup"><span data-stu-id="375c8-164">String</span></span>|<span data-ttu-id="375c8-165">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="375c8-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="375c8-166">要件</span><span class="sxs-lookup"><span data-stu-id="375c8-166">Requirements</span></span>

|<span data-ttu-id="375c8-167">要件</span><span class="sxs-lookup"><span data-stu-id="375c8-167">Requirement</span></span>| <span data-ttu-id="375c8-168">値</span><span class="sxs-lookup"><span data-stu-id="375c8-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="375c8-169">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="375c8-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="375c8-170">1.0</span><span class="sxs-lookup"><span data-stu-id="375c8-170">1.0</span></span>|
|[<span data-ttu-id="375c8-171">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="375c8-171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="375c8-172">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="375c8-172">Compose or Read</span></span>|
