---
title: Office 名前空間 - 要件セット 1.2
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: eff7896214866e71b92a1c8a0c72a16e622873f3
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067861"
---
# <a name="office"></a><span data-ttu-id="16c23-102">Office</span><span class="sxs-lookup"><span data-stu-id="16c23-102">Office</span></span>

<span data-ttu-id="16c23-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="16c23-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="16c23-105">要件</span><span class="sxs-lookup"><span data-stu-id="16c23-105">Requirements</span></span>

|<span data-ttu-id="16c23-106">要件</span><span class="sxs-lookup"><span data-stu-id="16c23-106">Requirement</span></span>| <span data-ttu-id="16c23-107">値</span><span class="sxs-lookup"><span data-stu-id="16c23-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="16c23-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="16c23-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="16c23-109">1.0</span><span class="sxs-lookup"><span data-stu-id="16c23-109">1.0</span></span>|
|[<span data-ttu-id="16c23-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="16c23-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="16c23-111">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="16c23-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="16c23-112">名前空間</span><span class="sxs-lookup"><span data-stu-id="16c23-112">Namespaces</span></span>

<span data-ttu-id="16c23-113">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="16c23-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="16c23-114">[MailboxEnums](/javascript/api/outlook_1_2/office.mailboxenums.attachmenttype): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="16c23-114">[MailboxEnums](/javascript/api/outlook_1_2/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="16c23-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="16c23-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="16c23-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="16c23-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="16c23-117">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="16c23-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="16c23-118">Type</span><span class="sxs-lookup"><span data-stu-id="16c23-118">Type</span></span>

*   <span data-ttu-id="16c23-119">String</span><span class="sxs-lookup"><span data-stu-id="16c23-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="16c23-120">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="16c23-120">Properties:</span></span>

|<span data-ttu-id="16c23-121">名前</span><span class="sxs-lookup"><span data-stu-id="16c23-121">Name</span></span>| <span data-ttu-id="16c23-122">型</span><span class="sxs-lookup"><span data-stu-id="16c23-122">Type</span></span>| <span data-ttu-id="16c23-123">説明</span><span class="sxs-lookup"><span data-stu-id="16c23-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="16c23-124">String</span><span class="sxs-lookup"><span data-stu-id="16c23-124">String</span></span>|<span data-ttu-id="16c23-125">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="16c23-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="16c23-126">String</span><span class="sxs-lookup"><span data-stu-id="16c23-126">String</span></span>|<span data-ttu-id="16c23-127">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="16c23-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="16c23-128">要件</span><span class="sxs-lookup"><span data-stu-id="16c23-128">Requirements</span></span>

|<span data-ttu-id="16c23-129">要件</span><span class="sxs-lookup"><span data-stu-id="16c23-129">Requirement</span></span>| <span data-ttu-id="16c23-130">値</span><span class="sxs-lookup"><span data-stu-id="16c23-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="16c23-131">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="16c23-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="16c23-132">1.0</span><span class="sxs-lookup"><span data-stu-id="16c23-132">1.0</span></span>|
|[<span data-ttu-id="16c23-133">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="16c23-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="16c23-134">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="16c23-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="16c23-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="16c23-135">CoercionType :String</span></span>

<span data-ttu-id="16c23-136">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="16c23-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="16c23-137">Type</span><span class="sxs-lookup"><span data-stu-id="16c23-137">Type</span></span>

*   <span data-ttu-id="16c23-138">String</span><span class="sxs-lookup"><span data-stu-id="16c23-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="16c23-139">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="16c23-139">Properties:</span></span>

|<span data-ttu-id="16c23-140">名前</span><span class="sxs-lookup"><span data-stu-id="16c23-140">Name</span></span>| <span data-ttu-id="16c23-141">型</span><span class="sxs-lookup"><span data-stu-id="16c23-141">Type</span></span>| <span data-ttu-id="16c23-142">説明</span><span class="sxs-lookup"><span data-stu-id="16c23-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="16c23-143">String</span><span class="sxs-lookup"><span data-stu-id="16c23-143">String</span></span>|<span data-ttu-id="16c23-144">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="16c23-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="16c23-145">String</span><span class="sxs-lookup"><span data-stu-id="16c23-145">String</span></span>|<span data-ttu-id="16c23-146">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="16c23-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="16c23-147">要件</span><span class="sxs-lookup"><span data-stu-id="16c23-147">Requirements</span></span>

|<span data-ttu-id="16c23-148">要件</span><span class="sxs-lookup"><span data-stu-id="16c23-148">Requirement</span></span>| <span data-ttu-id="16c23-149">値</span><span class="sxs-lookup"><span data-stu-id="16c23-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="16c23-150">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="16c23-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="16c23-151">1.0</span><span class="sxs-lookup"><span data-stu-id="16c23-151">1.0</span></span>|
|[<span data-ttu-id="16c23-152">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="16c23-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="16c23-153">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="16c23-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="16c23-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="16c23-154">SourceProperty :String</span></span>

<span data-ttu-id="16c23-155">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="16c23-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="16c23-156">Type</span><span class="sxs-lookup"><span data-stu-id="16c23-156">Type</span></span>

*   <span data-ttu-id="16c23-157">String</span><span class="sxs-lookup"><span data-stu-id="16c23-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="16c23-158">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="16c23-158">Properties:</span></span>

|<span data-ttu-id="16c23-159">名前</span><span class="sxs-lookup"><span data-stu-id="16c23-159">Name</span></span>| <span data-ttu-id="16c23-160">型</span><span class="sxs-lookup"><span data-stu-id="16c23-160">Type</span></span>| <span data-ttu-id="16c23-161">説明</span><span class="sxs-lookup"><span data-stu-id="16c23-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="16c23-162">String</span><span class="sxs-lookup"><span data-stu-id="16c23-162">String</span></span>|<span data-ttu-id="16c23-163">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="16c23-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="16c23-164">String</span><span class="sxs-lookup"><span data-stu-id="16c23-164">String</span></span>|<span data-ttu-id="16c23-165">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="16c23-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="16c23-166">要件</span><span class="sxs-lookup"><span data-stu-id="16c23-166">Requirements</span></span>

|<span data-ttu-id="16c23-167">要件</span><span class="sxs-lookup"><span data-stu-id="16c23-167">Requirement</span></span>| <span data-ttu-id="16c23-168">値</span><span class="sxs-lookup"><span data-stu-id="16c23-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="16c23-169">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="16c23-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="16c23-170">1.0</span><span class="sxs-lookup"><span data-stu-id="16c23-170">1.0</span></span>|
|[<span data-ttu-id="16c23-171">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="16c23-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="16c23-172">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="16c23-172">Compose or Read</span></span>|
