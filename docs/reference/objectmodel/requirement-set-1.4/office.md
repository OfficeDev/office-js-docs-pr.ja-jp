---
title: Office 名前空間-要件セット1.4
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: c60195ddfc42d962427127bf601bca3d41797566
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450325"
---
# <a name="office"></a><span data-ttu-id="e91ab-102">Office</span><span class="sxs-lookup"><span data-stu-id="e91ab-102">Office</span></span>

<span data-ttu-id="e91ab-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e91ab-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e91ab-105">要件</span><span class="sxs-lookup"><span data-stu-id="e91ab-105">Requirements</span></span>

|<span data-ttu-id="e91ab-106">要件</span><span class="sxs-lookup"><span data-stu-id="e91ab-106">Requirement</span></span>| <span data-ttu-id="e91ab-107">値</span><span class="sxs-lookup"><span data-stu-id="e91ab-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e91ab-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e91ab-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e91ab-109">1.0</span><span class="sxs-lookup"><span data-stu-id="e91ab-109">1.0</span></span>|
|[<span data-ttu-id="e91ab-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e91ab-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e91ab-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e91ab-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="e91ab-112">名前空間</span><span class="sxs-lookup"><span data-stu-id="e91ab-112">Namespaces</span></span>

<span data-ttu-id="e91ab-113">[context](Office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="e91ab-113">[context](Office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="e91ab-114">[MailboxEnums](/javascript/api/outlook_1_4/office.mailboxenums.attachmenttype): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="e91ab-114">[MailboxEnums](/javascript/api/outlook_1_4/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="e91ab-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="e91ab-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="e91ab-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="e91ab-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="e91ab-117">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="e91ab-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="e91ab-118">型</span><span class="sxs-lookup"><span data-stu-id="e91ab-118">Type</span></span>

*   <span data-ttu-id="e91ab-119">String</span><span class="sxs-lookup"><span data-stu-id="e91ab-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e91ab-120">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="e91ab-120">Properties:</span></span>

|<span data-ttu-id="e91ab-121">名前</span><span class="sxs-lookup"><span data-stu-id="e91ab-121">Name</span></span>| <span data-ttu-id="e91ab-122">種類</span><span class="sxs-lookup"><span data-stu-id="e91ab-122">Type</span></span>| <span data-ttu-id="e91ab-123">説明</span><span class="sxs-lookup"><span data-stu-id="e91ab-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="e91ab-124">String</span><span class="sxs-lookup"><span data-stu-id="e91ab-124">String</span></span>|<span data-ttu-id="e91ab-125">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="e91ab-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="e91ab-126">String</span><span class="sxs-lookup"><span data-stu-id="e91ab-126">String</span></span>|<span data-ttu-id="e91ab-127">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="e91ab-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e91ab-128">要件</span><span class="sxs-lookup"><span data-stu-id="e91ab-128">Requirements</span></span>

|<span data-ttu-id="e91ab-129">要件</span><span class="sxs-lookup"><span data-stu-id="e91ab-129">Requirement</span></span>| <span data-ttu-id="e91ab-130">値</span><span class="sxs-lookup"><span data-stu-id="e91ab-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="e91ab-131">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e91ab-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e91ab-132">1.0</span><span class="sxs-lookup"><span data-stu-id="e91ab-132">1.0</span></span>|
|[<span data-ttu-id="e91ab-133">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e91ab-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e91ab-134">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e91ab-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="e91ab-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="e91ab-135">CoercionType :String</span></span>

<span data-ttu-id="e91ab-136">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="e91ab-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e91ab-137">型</span><span class="sxs-lookup"><span data-stu-id="e91ab-137">Type</span></span>

*   <span data-ttu-id="e91ab-138">String</span><span class="sxs-lookup"><span data-stu-id="e91ab-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e91ab-139">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="e91ab-139">Properties:</span></span>

|<span data-ttu-id="e91ab-140">名前</span><span class="sxs-lookup"><span data-stu-id="e91ab-140">Name</span></span>| <span data-ttu-id="e91ab-141">種類</span><span class="sxs-lookup"><span data-stu-id="e91ab-141">Type</span></span>| <span data-ttu-id="e91ab-142">説明</span><span class="sxs-lookup"><span data-stu-id="e91ab-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="e91ab-143">String</span><span class="sxs-lookup"><span data-stu-id="e91ab-143">String</span></span>|<span data-ttu-id="e91ab-144">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="e91ab-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="e91ab-145">String</span><span class="sxs-lookup"><span data-stu-id="e91ab-145">String</span></span>|<span data-ttu-id="e91ab-146">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="e91ab-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e91ab-147">要件</span><span class="sxs-lookup"><span data-stu-id="e91ab-147">Requirements</span></span>

|<span data-ttu-id="e91ab-148">要件</span><span class="sxs-lookup"><span data-stu-id="e91ab-148">Requirement</span></span>| <span data-ttu-id="e91ab-149">値</span><span class="sxs-lookup"><span data-stu-id="e91ab-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="e91ab-150">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e91ab-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e91ab-151">1.0</span><span class="sxs-lookup"><span data-stu-id="e91ab-151">1.0</span></span>|
|[<span data-ttu-id="e91ab-152">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e91ab-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e91ab-153">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e91ab-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="e91ab-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="e91ab-154">SourceProperty :String</span></span>

<span data-ttu-id="e91ab-155">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="e91ab-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e91ab-156">型</span><span class="sxs-lookup"><span data-stu-id="e91ab-156">Type</span></span>

*   <span data-ttu-id="e91ab-157">String</span><span class="sxs-lookup"><span data-stu-id="e91ab-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e91ab-158">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="e91ab-158">Properties:</span></span>

|<span data-ttu-id="e91ab-159">名前</span><span class="sxs-lookup"><span data-stu-id="e91ab-159">Name</span></span>| <span data-ttu-id="e91ab-160">種類</span><span class="sxs-lookup"><span data-stu-id="e91ab-160">Type</span></span>| <span data-ttu-id="e91ab-161">説明</span><span class="sxs-lookup"><span data-stu-id="e91ab-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="e91ab-162">String</span><span class="sxs-lookup"><span data-stu-id="e91ab-162">String</span></span>|<span data-ttu-id="e91ab-163">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="e91ab-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="e91ab-164">String</span><span class="sxs-lookup"><span data-stu-id="e91ab-164">String</span></span>|<span data-ttu-id="e91ab-165">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="e91ab-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e91ab-166">要件</span><span class="sxs-lookup"><span data-stu-id="e91ab-166">Requirements</span></span>

|<span data-ttu-id="e91ab-167">要件</span><span class="sxs-lookup"><span data-stu-id="e91ab-167">Requirement</span></span>| <span data-ttu-id="e91ab-168">値</span><span class="sxs-lookup"><span data-stu-id="e91ab-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="e91ab-169">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e91ab-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e91ab-170">1.0</span><span class="sxs-lookup"><span data-stu-id="e91ab-170">1.0</span></span>|
|[<span data-ttu-id="e91ab-171">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e91ab-171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e91ab-172">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e91ab-172">Compose or Read</span></span>|
