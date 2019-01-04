---
title: Office 名前空間 - 要件セット 1.1
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: af2a48d5bc943d4f443c32777fefaf8ed4a30032
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457433"
---
# <a name="office"></a><span data-ttu-id="31685-102">Office</span><span class="sxs-lookup"><span data-stu-id="31685-102">Office</span></span>

<span data-ttu-id="31685-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="31685-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="31685-105">要件</span><span class="sxs-lookup"><span data-stu-id="31685-105">Requirements</span></span>

|<span data-ttu-id="31685-106">要件</span><span class="sxs-lookup"><span data-stu-id="31685-106">Requirement</span></span>| <span data-ttu-id="31685-107">値</span><span class="sxs-lookup"><span data-stu-id="31685-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="31685-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="31685-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31685-109">1.0</span><span class="sxs-lookup"><span data-stu-id="31685-109">1.0</span></span>|
|[<span data-ttu-id="31685-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="31685-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31685-111">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="31685-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="31685-112">名前空間</span><span class="sxs-lookup"><span data-stu-id="31685-112">Namespaces</span></span>

<span data-ttu-id="31685-113">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="31685-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="31685-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="31685-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="31685-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="31685-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="31685-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="31685-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="31685-117">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="31685-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="31685-118">型:</span><span class="sxs-lookup"><span data-stu-id="31685-118">Type:</span></span>

*   <span data-ttu-id="31685-119">String</span><span class="sxs-lookup"><span data-stu-id="31685-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="31685-120">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="31685-120">Properties:</span></span>

|<span data-ttu-id="31685-121">名前</span><span class="sxs-lookup"><span data-stu-id="31685-121">Name</span></span>| <span data-ttu-id="31685-122">型</span><span class="sxs-lookup"><span data-stu-id="31685-122">Type</span></span>| <span data-ttu-id="31685-123">説明</span><span class="sxs-lookup"><span data-stu-id="31685-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="31685-124">String</span><span class="sxs-lookup"><span data-stu-id="31685-124">String</span></span>|<span data-ttu-id="31685-125">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="31685-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="31685-126">String</span><span class="sxs-lookup"><span data-stu-id="31685-126">String</span></span>|<span data-ttu-id="31685-127">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="31685-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="31685-128">要件</span><span class="sxs-lookup"><span data-stu-id="31685-128">Requirements</span></span>

|<span data-ttu-id="31685-129">要件</span><span class="sxs-lookup"><span data-stu-id="31685-129">Requirement</span></span>| <span data-ttu-id="31685-130">値</span><span class="sxs-lookup"><span data-stu-id="31685-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="31685-131">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="31685-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31685-132">1.0</span><span class="sxs-lookup"><span data-stu-id="31685-132">1.0</span></span>|
|[<span data-ttu-id="31685-133">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="31685-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31685-134">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="31685-134">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="31685-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="31685-135">CoercionType :String</span></span>

<span data-ttu-id="31685-136">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="31685-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="31685-137">型:</span><span class="sxs-lookup"><span data-stu-id="31685-137">Type:</span></span>

*   <span data-ttu-id="31685-138">String</span><span class="sxs-lookup"><span data-stu-id="31685-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="31685-139">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="31685-139">Properties:</span></span>

|<span data-ttu-id="31685-140">名前</span><span class="sxs-lookup"><span data-stu-id="31685-140">Name</span></span>| <span data-ttu-id="31685-141">型</span><span class="sxs-lookup"><span data-stu-id="31685-141">Type</span></span>| <span data-ttu-id="31685-142">説明</span><span class="sxs-lookup"><span data-stu-id="31685-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="31685-143">String</span><span class="sxs-lookup"><span data-stu-id="31685-143">String</span></span>|<span data-ttu-id="31685-144">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="31685-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="31685-145">String</span><span class="sxs-lookup"><span data-stu-id="31685-145">String</span></span>|<span data-ttu-id="31685-146">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="31685-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="31685-147">要件</span><span class="sxs-lookup"><span data-stu-id="31685-147">Requirements</span></span>

|<span data-ttu-id="31685-148">要件</span><span class="sxs-lookup"><span data-stu-id="31685-148">Requirement</span></span>| <span data-ttu-id="31685-149">値</span><span class="sxs-lookup"><span data-stu-id="31685-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="31685-150">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="31685-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31685-151">1.0</span><span class="sxs-lookup"><span data-stu-id="31685-151">1.0</span></span>|
|[<span data-ttu-id="31685-152">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="31685-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31685-153">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="31685-153">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="31685-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="31685-154">SourceProperty :String</span></span>

<span data-ttu-id="31685-155">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="31685-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="31685-156">型:</span><span class="sxs-lookup"><span data-stu-id="31685-156">Type:</span></span>

*   <span data-ttu-id="31685-157">String</span><span class="sxs-lookup"><span data-stu-id="31685-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="31685-158">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="31685-158">Properties:</span></span>

|<span data-ttu-id="31685-159">名前</span><span class="sxs-lookup"><span data-stu-id="31685-159">Name</span></span>| <span data-ttu-id="31685-160">型</span><span class="sxs-lookup"><span data-stu-id="31685-160">Type</span></span>| <span data-ttu-id="31685-161">説明</span><span class="sxs-lookup"><span data-stu-id="31685-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="31685-162">String</span><span class="sxs-lookup"><span data-stu-id="31685-162">String</span></span>|<span data-ttu-id="31685-163">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="31685-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="31685-164">String</span><span class="sxs-lookup"><span data-stu-id="31685-164">String</span></span>|<span data-ttu-id="31685-165">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="31685-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="31685-166">要件</span><span class="sxs-lookup"><span data-stu-id="31685-166">Requirements</span></span>

|<span data-ttu-id="31685-167">要件</span><span class="sxs-lookup"><span data-stu-id="31685-167">Requirement</span></span>| <span data-ttu-id="31685-168">値</span><span class="sxs-lookup"><span data-stu-id="31685-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="31685-169">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="31685-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31685-170">1.0</span><span class="sxs-lookup"><span data-stu-id="31685-170">1.0</span></span>|
|[<span data-ttu-id="31685-171">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="31685-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31685-172">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="31685-172">Compose or read</span></span>|