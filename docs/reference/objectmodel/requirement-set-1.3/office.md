 

# <a name="office"></a><span data-ttu-id="4e018-101">Office</span><span class="sxs-lookup"><span data-stu-id="4e018-101">Office</span></span>

<span data-ttu-id="4e018-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共有 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4e018-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e018-104">要件</span><span class="sxs-lookup"><span data-stu-id="4e018-104">Requirements</span></span>

|<span data-ttu-id="4e018-105">要件</span><span class="sxs-lookup"><span data-stu-id="4e018-105">Requirement</span></span>| <span data-ttu-id="4e018-106">値</span><span class="sxs-lookup"><span data-stu-id="4e018-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e018-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4e018-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e018-108">1.0</span><span class="sxs-lookup"><span data-stu-id="4e018-108">1.0</span></span>|
|[<span data-ttu-id="4e018-109">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4e018-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4e018-110">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="4e018-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="4e018-111">名前空間</span><span class="sxs-lookup"><span data-stu-id="4e018-111">Namespaces</span></span>

<span data-ttu-id="4e018-112">[context](office.context.md):Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="4e018-112">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="4e018-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype):ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="4e018-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="4e018-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="4e018-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="4e018-115">AsyncResultStatus: 文字列</span><span class="sxs-lookup"><span data-stu-id="4e018-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="4e018-116">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="4e018-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="4e018-117">型:</span><span class="sxs-lookup"><span data-stu-id="4e018-117">Type:</span></span>

*   <span data-ttu-id="4e018-118">文字列</span><span class="sxs-lookup"><span data-stu-id="4e018-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4e018-119">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="4e018-119">Properties:</span></span>

|<span data-ttu-id="4e018-120">名前</span><span class="sxs-lookup"><span data-stu-id="4e018-120">Name</span></span>| <span data-ttu-id="4e018-121">型</span><span class="sxs-lookup"><span data-stu-id="4e018-121">Type</span></span>| <span data-ttu-id="4e018-122">説明</span><span class="sxs-lookup"><span data-stu-id="4e018-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="4e018-123">文字列</span><span class="sxs-lookup"><span data-stu-id="4e018-123">String</span></span>|<span data-ttu-id="4e018-124">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="4e018-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="4e018-125">文字列</span><span class="sxs-lookup"><span data-stu-id="4e018-125">String</span></span>|<span data-ttu-id="4e018-126">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="4e018-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e018-127">要件</span><span class="sxs-lookup"><span data-stu-id="4e018-127">Requirements</span></span>

|<span data-ttu-id="4e018-128">要件</span><span class="sxs-lookup"><span data-stu-id="4e018-128">Requirement</span></span>| <span data-ttu-id="4e018-129">値</span><span class="sxs-lookup"><span data-stu-id="4e018-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e018-130">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4e018-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e018-131">1.0</span><span class="sxs-lookup"><span data-stu-id="4e018-131">1.0</span></span>|
|[<span data-ttu-id="4e018-132">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4e018-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4e018-133">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="4e018-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="4e018-134">CoercionType: 文字列</span><span class="sxs-lookup"><span data-stu-id="4e018-134">CoercionType :String</span></span>

<span data-ttu-id="4e018-135">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="4e018-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4e018-136">型:</span><span class="sxs-lookup"><span data-stu-id="4e018-136">Type:</span></span>

*   <span data-ttu-id="4e018-137">文字列</span><span class="sxs-lookup"><span data-stu-id="4e018-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4e018-138">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="4e018-138">Properties:</span></span>

|<span data-ttu-id="4e018-139">名前</span><span class="sxs-lookup"><span data-stu-id="4e018-139">Name</span></span>| <span data-ttu-id="4e018-140">型</span><span class="sxs-lookup"><span data-stu-id="4e018-140">Type</span></span>| <span data-ttu-id="4e018-141">説明</span><span class="sxs-lookup"><span data-stu-id="4e018-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="4e018-142">文字列</span><span class="sxs-lookup"><span data-stu-id="4e018-142">String</span></span>|<span data-ttu-id="4e018-143">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="4e018-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="4e018-144">文字列</span><span class="sxs-lookup"><span data-stu-id="4e018-144">String</span></span>|<span data-ttu-id="4e018-145">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="4e018-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e018-146">要件</span><span class="sxs-lookup"><span data-stu-id="4e018-146">Requirements</span></span>

|<span data-ttu-id="4e018-147">要件</span><span class="sxs-lookup"><span data-stu-id="4e018-147">Requirement</span></span>| <span data-ttu-id="4e018-148">値</span><span class="sxs-lookup"><span data-stu-id="4e018-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e018-149">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4e018-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e018-150">1.0</span><span class="sxs-lookup"><span data-stu-id="4e018-150">1.0</span></span>|
|[<span data-ttu-id="4e018-151">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4e018-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4e018-152">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="4e018-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="4e018-153">SourceProperty: 文字列</span><span class="sxs-lookup"><span data-stu-id="4e018-153">SourceProperty :String</span></span>

<span data-ttu-id="4e018-154">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="4e018-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4e018-155">型:</span><span class="sxs-lookup"><span data-stu-id="4e018-155">Type:</span></span>

*   <span data-ttu-id="4e018-156">文字列</span><span class="sxs-lookup"><span data-stu-id="4e018-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4e018-157">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="4e018-157">Properties:</span></span>

|<span data-ttu-id="4e018-158">名前</span><span class="sxs-lookup"><span data-stu-id="4e018-158">Name</span></span>| <span data-ttu-id="4e018-159">型</span><span class="sxs-lookup"><span data-stu-id="4e018-159">Type</span></span>| <span data-ttu-id="4e018-160">説明</span><span class="sxs-lookup"><span data-stu-id="4e018-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="4e018-161">文字列</span><span class="sxs-lookup"><span data-stu-id="4e018-161">String</span></span>|<span data-ttu-id="4e018-162">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="4e018-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="4e018-163">文字列</span><span class="sxs-lookup"><span data-stu-id="4e018-163">String</span></span>|<span data-ttu-id="4e018-164">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="4e018-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e018-165">要件</span><span class="sxs-lookup"><span data-stu-id="4e018-165">Requirements</span></span>

|<span data-ttu-id="4e018-166">要件</span><span class="sxs-lookup"><span data-stu-id="4e018-166">Requirement</span></span>| <span data-ttu-id="4e018-167">値</span><span class="sxs-lookup"><span data-stu-id="4e018-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e018-168">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4e018-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e018-169">1.0</span><span class="sxs-lookup"><span data-stu-id="4e018-169">1.0</span></span>|
|[<span data-ttu-id="4e018-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4e018-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4e018-171">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="4e018-171">Compose or read</span></span>|