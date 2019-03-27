---
title: Office 名前空間 - 要件セット 1.1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: eda5e1fb5f2c11ae91e4a1479892c36ec23e1897
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871998"
---
# <a name="office"></a><span data-ttu-id="f0a8c-102">Office</span><span class="sxs-lookup"><span data-stu-id="f0a8c-102">Office</span></span>

<span data-ttu-id="f0a8c-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f0a8c-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0a8c-105">要件</span><span class="sxs-lookup"><span data-stu-id="f0a8c-105">Requirements</span></span>

|<span data-ttu-id="f0a8c-106">要件</span><span class="sxs-lookup"><span data-stu-id="f0a8c-106">Requirement</span></span>| <span data-ttu-id="f0a8c-107">値</span><span class="sxs-lookup"><span data-stu-id="f0a8c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0a8c-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0a8c-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0a8c-109">1.0</span><span class="sxs-lookup"><span data-stu-id="f0a8c-109">1.0</span></span>|
|[<span data-ttu-id="f0a8c-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0a8c-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0a8c-111">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0a8c-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="f0a8c-112">名前空間</span><span class="sxs-lookup"><span data-stu-id="f0a8c-112">Namespaces</span></span>

<span data-ttu-id="f0a8c-113">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f0a8c-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="f0a8c-114">[MailboxEnums](/javascript/api/outlook_1_1/office.mailboxenums.attachmenttype): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="f0a8c-114">[MailboxEnums](/javascript/api/outlook_1_1/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="f0a8c-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0a8c-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="f0a8c-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="f0a8c-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="f0a8c-117">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="f0a8c-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f0a8c-118">型</span><span class="sxs-lookup"><span data-stu-id="f0a8c-118">Type</span></span>

*   <span data-ttu-id="f0a8c-119">String</span><span class="sxs-lookup"><span data-stu-id="f0a8c-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f0a8c-120">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="f0a8c-120">Properties:</span></span>

|<span data-ttu-id="f0a8c-121">名前</span><span class="sxs-lookup"><span data-stu-id="f0a8c-121">Name</span></span>| <span data-ttu-id="f0a8c-122">種類</span><span class="sxs-lookup"><span data-stu-id="f0a8c-122">Type</span></span>| <span data-ttu-id="f0a8c-123">説明</span><span class="sxs-lookup"><span data-stu-id="f0a8c-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f0a8c-124">String</span><span class="sxs-lookup"><span data-stu-id="f0a8c-124">String</span></span>|<span data-ttu-id="f0a8c-125">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="f0a8c-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f0a8c-126">String</span><span class="sxs-lookup"><span data-stu-id="f0a8c-126">String</span></span>|<span data-ttu-id="f0a8c-127">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="f0a8c-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f0a8c-128">要件</span><span class="sxs-lookup"><span data-stu-id="f0a8c-128">Requirements</span></span>

|<span data-ttu-id="f0a8c-129">要件</span><span class="sxs-lookup"><span data-stu-id="f0a8c-129">Requirement</span></span>| <span data-ttu-id="f0a8c-130">値</span><span class="sxs-lookup"><span data-stu-id="f0a8c-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0a8c-131">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0a8c-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0a8c-132">1.0</span><span class="sxs-lookup"><span data-stu-id="f0a8c-132">1.0</span></span>|
|[<span data-ttu-id="f0a8c-133">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0a8c-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0a8c-134">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0a8c-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="f0a8c-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="f0a8c-135">CoercionType :String</span></span>

<span data-ttu-id="f0a8c-136">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="f0a8c-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f0a8c-137">型</span><span class="sxs-lookup"><span data-stu-id="f0a8c-137">Type</span></span>

*   <span data-ttu-id="f0a8c-138">String</span><span class="sxs-lookup"><span data-stu-id="f0a8c-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f0a8c-139">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="f0a8c-139">Properties:</span></span>

|<span data-ttu-id="f0a8c-140">名前</span><span class="sxs-lookup"><span data-stu-id="f0a8c-140">Name</span></span>| <span data-ttu-id="f0a8c-141">種類</span><span class="sxs-lookup"><span data-stu-id="f0a8c-141">Type</span></span>| <span data-ttu-id="f0a8c-142">説明</span><span class="sxs-lookup"><span data-stu-id="f0a8c-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f0a8c-143">String</span><span class="sxs-lookup"><span data-stu-id="f0a8c-143">String</span></span>|<span data-ttu-id="f0a8c-144">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="f0a8c-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f0a8c-145">String</span><span class="sxs-lookup"><span data-stu-id="f0a8c-145">String</span></span>|<span data-ttu-id="f0a8c-146">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="f0a8c-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f0a8c-147">要件</span><span class="sxs-lookup"><span data-stu-id="f0a8c-147">Requirements</span></span>

|<span data-ttu-id="f0a8c-148">要件</span><span class="sxs-lookup"><span data-stu-id="f0a8c-148">Requirement</span></span>| <span data-ttu-id="f0a8c-149">値</span><span class="sxs-lookup"><span data-stu-id="f0a8c-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0a8c-150">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0a8c-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0a8c-151">1.0</span><span class="sxs-lookup"><span data-stu-id="f0a8c-151">1.0</span></span>|
|[<span data-ttu-id="f0a8c-152">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0a8c-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0a8c-153">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0a8c-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="f0a8c-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="f0a8c-154">SourceProperty :String</span></span>

<span data-ttu-id="f0a8c-155">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="f0a8c-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f0a8c-156">型</span><span class="sxs-lookup"><span data-stu-id="f0a8c-156">Type</span></span>

*   <span data-ttu-id="f0a8c-157">String</span><span class="sxs-lookup"><span data-stu-id="f0a8c-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f0a8c-158">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="f0a8c-158">Properties:</span></span>

|<span data-ttu-id="f0a8c-159">名前</span><span class="sxs-lookup"><span data-stu-id="f0a8c-159">Name</span></span>| <span data-ttu-id="f0a8c-160">種類</span><span class="sxs-lookup"><span data-stu-id="f0a8c-160">Type</span></span>| <span data-ttu-id="f0a8c-161">説明</span><span class="sxs-lookup"><span data-stu-id="f0a8c-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f0a8c-162">String</span><span class="sxs-lookup"><span data-stu-id="f0a8c-162">String</span></span>|<span data-ttu-id="f0a8c-163">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="f0a8c-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f0a8c-164">String</span><span class="sxs-lookup"><span data-stu-id="f0a8c-164">String</span></span>|<span data-ttu-id="f0a8c-165">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="f0a8c-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f0a8c-166">要件</span><span class="sxs-lookup"><span data-stu-id="f0a8c-166">Requirements</span></span>

|<span data-ttu-id="f0a8c-167">要件</span><span class="sxs-lookup"><span data-stu-id="f0a8c-167">Requirement</span></span>| <span data-ttu-id="f0a8c-168">値</span><span class="sxs-lookup"><span data-stu-id="f0a8c-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0a8c-169">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0a8c-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0a8c-170">1.0</span><span class="sxs-lookup"><span data-stu-id="f0a8c-170">1.0</span></span>|
|[<span data-ttu-id="f0a8c-171">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0a8c-171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0a8c-172">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0a8c-172">Compose or Read</span></span>|
