---
title: Office 名前空間 - 要件セット 1.4
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: f447c3c46db553c0a6d4ba8bf2d4e108848c3f4d
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457685"
---
# <a name="office"></a><span data-ttu-id="88b6f-102">Office</span><span class="sxs-lookup"><span data-stu-id="88b6f-102">Office</span></span>

<span data-ttu-id="88b6f-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共通 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="88b6f-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="88b6f-105">要件</span><span class="sxs-lookup"><span data-stu-id="88b6f-105">Requirements</span></span>

|<span data-ttu-id="88b6f-106">要件</span><span class="sxs-lookup"><span data-stu-id="88b6f-106">Requirement</span></span>| <span data-ttu-id="88b6f-107">値</span><span class="sxs-lookup"><span data-stu-id="88b6f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="88b6f-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88b6f-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88b6f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="88b6f-109">1.0</span></span>|
|[<span data-ttu-id="88b6f-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88b6f-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88b6f-111">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="88b6f-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="88b6f-112">名前空間</span><span class="sxs-lookup"><span data-stu-id="88b6f-112">Namespaces</span></span>

<span data-ttu-id="88b6f-113">[context](Office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="88b6f-113">[context](Office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="88b6f-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="88b6f-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="88b6f-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="88b6f-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="88b6f-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="88b6f-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="88b6f-117">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="88b6f-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="88b6f-118">型:</span><span class="sxs-lookup"><span data-stu-id="88b6f-118">Type:</span></span>

*   <span data-ttu-id="88b6f-119">String</span><span class="sxs-lookup"><span data-stu-id="88b6f-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="88b6f-120">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="88b6f-120">Properties:</span></span>

|<span data-ttu-id="88b6f-121">名前</span><span class="sxs-lookup"><span data-stu-id="88b6f-121">Name</span></span>| <span data-ttu-id="88b6f-122">型</span><span class="sxs-lookup"><span data-stu-id="88b6f-122">Type</span></span>| <span data-ttu-id="88b6f-123">説明</span><span class="sxs-lookup"><span data-stu-id="88b6f-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="88b6f-124">String</span><span class="sxs-lookup"><span data-stu-id="88b6f-124">String</span></span>|<span data-ttu-id="88b6f-125">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="88b6f-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="88b6f-126">String</span><span class="sxs-lookup"><span data-stu-id="88b6f-126">String</span></span>|<span data-ttu-id="88b6f-127">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="88b6f-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="88b6f-128">要件</span><span class="sxs-lookup"><span data-stu-id="88b6f-128">Requirements</span></span>

|<span data-ttu-id="88b6f-129">要件</span><span class="sxs-lookup"><span data-stu-id="88b6f-129">Requirement</span></span>| <span data-ttu-id="88b6f-130">値</span><span class="sxs-lookup"><span data-stu-id="88b6f-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="88b6f-131">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88b6f-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88b6f-132">1.0</span><span class="sxs-lookup"><span data-stu-id="88b6f-132">1.0</span></span>|
|[<span data-ttu-id="88b6f-133">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88b6f-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88b6f-134">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="88b6f-134">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="88b6f-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="88b6f-135">CoercionType :String</span></span>

<span data-ttu-id="88b6f-136">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="88b6f-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="88b6f-137">型:</span><span class="sxs-lookup"><span data-stu-id="88b6f-137">Type:</span></span>

*   <span data-ttu-id="88b6f-138">String</span><span class="sxs-lookup"><span data-stu-id="88b6f-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="88b6f-139">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="88b6f-139">Properties:</span></span>

|<span data-ttu-id="88b6f-140">名前</span><span class="sxs-lookup"><span data-stu-id="88b6f-140">Name</span></span>| <span data-ttu-id="88b6f-141">型</span><span class="sxs-lookup"><span data-stu-id="88b6f-141">Type</span></span>| <span data-ttu-id="88b6f-142">説明</span><span class="sxs-lookup"><span data-stu-id="88b6f-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="88b6f-143">String</span><span class="sxs-lookup"><span data-stu-id="88b6f-143">String</span></span>|<span data-ttu-id="88b6f-144">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="88b6f-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="88b6f-145">String</span><span class="sxs-lookup"><span data-stu-id="88b6f-145">String</span></span>|<span data-ttu-id="88b6f-146">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="88b6f-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="88b6f-147">要件</span><span class="sxs-lookup"><span data-stu-id="88b6f-147">Requirements</span></span>

|<span data-ttu-id="88b6f-148">要件</span><span class="sxs-lookup"><span data-stu-id="88b6f-148">Requirement</span></span>| <span data-ttu-id="88b6f-149">値</span><span class="sxs-lookup"><span data-stu-id="88b6f-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="88b6f-150">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88b6f-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88b6f-151">1.0</span><span class="sxs-lookup"><span data-stu-id="88b6f-151">1.0</span></span>|
|[<span data-ttu-id="88b6f-152">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88b6f-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88b6f-153">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="88b6f-153">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="88b6f-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="88b6f-154">SourceProperty :String</span></span>

<span data-ttu-id="88b6f-155">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="88b6f-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="88b6f-156">型:</span><span class="sxs-lookup"><span data-stu-id="88b6f-156">Type:</span></span>

*   <span data-ttu-id="88b6f-157">String</span><span class="sxs-lookup"><span data-stu-id="88b6f-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="88b6f-158">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="88b6f-158">Properties:</span></span>

|<span data-ttu-id="88b6f-159">名前</span><span class="sxs-lookup"><span data-stu-id="88b6f-159">Name</span></span>| <span data-ttu-id="88b6f-160">型</span><span class="sxs-lookup"><span data-stu-id="88b6f-160">Type</span></span>| <span data-ttu-id="88b6f-161">説明</span><span class="sxs-lookup"><span data-stu-id="88b6f-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="88b6f-162">String</span><span class="sxs-lookup"><span data-stu-id="88b6f-162">String</span></span>|<span data-ttu-id="88b6f-163">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="88b6f-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="88b6f-164">String</span><span class="sxs-lookup"><span data-stu-id="88b6f-164">String</span></span>|<span data-ttu-id="88b6f-165">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="88b6f-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="88b6f-166">要件</span><span class="sxs-lookup"><span data-stu-id="88b6f-166">Requirements</span></span>

|<span data-ttu-id="88b6f-167">要件</span><span class="sxs-lookup"><span data-stu-id="88b6f-167">Requirement</span></span>| <span data-ttu-id="88b6f-168">値</span><span class="sxs-lookup"><span data-stu-id="88b6f-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="88b6f-169">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88b6f-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88b6f-170">1.0</span><span class="sxs-lookup"><span data-stu-id="88b6f-170">1.0</span></span>|
|[<span data-ttu-id="88b6f-171">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88b6f-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88b6f-172">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="88b6f-172">Compose or read</span></span>|