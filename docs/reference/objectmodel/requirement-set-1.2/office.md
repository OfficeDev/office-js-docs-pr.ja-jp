 

# <a name="office"></a><span data-ttu-id="50b13-101">Office</span><span class="sxs-lookup"><span data-stu-id="50b13-101">Office</span></span>

<span data-ttu-id="50b13-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共有 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="50b13-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="50b13-104">要件</span><span class="sxs-lookup"><span data-stu-id="50b13-104">Requirements</span></span>

|<span data-ttu-id="50b13-105">要件</span><span class="sxs-lookup"><span data-stu-id="50b13-105">Requirement</span></span>| <span data-ttu-id="50b13-106">値</span><span class="sxs-lookup"><span data-stu-id="50b13-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="50b13-107">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="50b13-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="50b13-108">1.0</span><span class="sxs-lookup"><span data-stu-id="50b13-108">1.0</span></span>|
|[<span data-ttu-id="50b13-109">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="50b13-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="50b13-110">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="50b13-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="50b13-111">名前空間</span><span class="sxs-lookup"><span data-stu-id="50b13-111">Namespaces</span></span>

<span data-ttu-id="50b13-112">[context](office.context.md):Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="50b13-112">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="50b13-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype):ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="50b13-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="50b13-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="50b13-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="50b13-115">AsyncResultStatus: 文字列</span><span class="sxs-lookup"><span data-stu-id="50b13-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="50b13-116">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="50b13-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="50b13-117">種類:</span><span class="sxs-lookup"><span data-stu-id="50b13-117">Type:</span></span>

*   <span data-ttu-id="50b13-118">文字列</span><span class="sxs-lookup"><span data-stu-id="50b13-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="50b13-119">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="50b13-119">Properties:</span></span>

|<span data-ttu-id="50b13-120">名前</span><span class="sxs-lookup"><span data-stu-id="50b13-120">Name</span></span>| <span data-ttu-id="50b13-121">種類</span><span class="sxs-lookup"><span data-stu-id="50b13-121">Type</span></span>| <span data-ttu-id="50b13-122">説明</span><span class="sxs-lookup"><span data-stu-id="50b13-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="50b13-123">文字列</span><span class="sxs-lookup"><span data-stu-id="50b13-123">String</span></span>|<span data-ttu-id="50b13-124">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="50b13-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="50b13-125">文字列</span><span class="sxs-lookup"><span data-stu-id="50b13-125">String</span></span>|<span data-ttu-id="50b13-126">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="50b13-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="50b13-127">要件</span><span class="sxs-lookup"><span data-stu-id="50b13-127">Requirements</span></span>

|<span data-ttu-id="50b13-128">要件</span><span class="sxs-lookup"><span data-stu-id="50b13-128">Requirement</span></span>| <span data-ttu-id="50b13-129">値</span><span class="sxs-lookup"><span data-stu-id="50b13-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="50b13-130">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="50b13-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="50b13-131">1.0</span><span class="sxs-lookup"><span data-stu-id="50b13-131">1.0</span></span>|
|[<span data-ttu-id="50b13-132">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="50b13-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="50b13-133">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="50b13-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="50b13-134">CoercionType: 文字列</span><span class="sxs-lookup"><span data-stu-id="50b13-134">CoercionType :String</span></span>

<span data-ttu-id="50b13-135">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="50b13-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="50b13-136">種類:</span><span class="sxs-lookup"><span data-stu-id="50b13-136">Type:</span></span>

*   <span data-ttu-id="50b13-137">文字列</span><span class="sxs-lookup"><span data-stu-id="50b13-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="50b13-138">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="50b13-138">Properties:</span></span>

|<span data-ttu-id="50b13-139">名前</span><span class="sxs-lookup"><span data-stu-id="50b13-139">Name</span></span>| <span data-ttu-id="50b13-140">種類</span><span class="sxs-lookup"><span data-stu-id="50b13-140">Type</span></span>| <span data-ttu-id="50b13-141">説明</span><span class="sxs-lookup"><span data-stu-id="50b13-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="50b13-142">文字列</span><span class="sxs-lookup"><span data-stu-id="50b13-142">String</span></span>|<span data-ttu-id="50b13-143">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="50b13-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="50b13-144">文字列</span><span class="sxs-lookup"><span data-stu-id="50b13-144">String</span></span>|<span data-ttu-id="50b13-145">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="50b13-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="50b13-146">要件</span><span class="sxs-lookup"><span data-stu-id="50b13-146">Requirements</span></span>

|<span data-ttu-id="50b13-147">要件</span><span class="sxs-lookup"><span data-stu-id="50b13-147">Requirement</span></span>| <span data-ttu-id="50b13-148">値</span><span class="sxs-lookup"><span data-stu-id="50b13-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="50b13-149">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="50b13-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="50b13-150">1.0</span><span class="sxs-lookup"><span data-stu-id="50b13-150">1.0</span></span>|
|[<span data-ttu-id="50b13-151">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="50b13-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="50b13-152">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="50b13-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="50b13-153">SourceProperty: 文字列</span><span class="sxs-lookup"><span data-stu-id="50b13-153">SourceProperty :String</span></span>

<span data-ttu-id="50b13-154">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="50b13-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="50b13-155">種類:</span><span class="sxs-lookup"><span data-stu-id="50b13-155">Type:</span></span>

*   <span data-ttu-id="50b13-156">文字列</span><span class="sxs-lookup"><span data-stu-id="50b13-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="50b13-157">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="50b13-157">Properties:</span></span>

|<span data-ttu-id="50b13-158">名前</span><span class="sxs-lookup"><span data-stu-id="50b13-158">Name</span></span>| <span data-ttu-id="50b13-159">種類</span><span class="sxs-lookup"><span data-stu-id="50b13-159">Type</span></span>| <span data-ttu-id="50b13-160">説明</span><span class="sxs-lookup"><span data-stu-id="50b13-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="50b13-161">文字列</span><span class="sxs-lookup"><span data-stu-id="50b13-161">String</span></span>|<span data-ttu-id="50b13-162">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="50b13-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="50b13-163">文字列</span><span class="sxs-lookup"><span data-stu-id="50b13-163">String</span></span>|<span data-ttu-id="50b13-164">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="50b13-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="50b13-165">要件</span><span class="sxs-lookup"><span data-stu-id="50b13-165">Requirements</span></span>

|<span data-ttu-id="50b13-166">要件</span><span class="sxs-lookup"><span data-stu-id="50b13-166">Requirement</span></span>| <span data-ttu-id="50b13-167">値</span><span class="sxs-lookup"><span data-stu-id="50b13-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="50b13-168">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="50b13-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="50b13-169">1.0</span><span class="sxs-lookup"><span data-stu-id="50b13-169">1.0</span></span>|
|[<span data-ttu-id="50b13-170">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="50b13-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="50b13-171">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="50b13-171">Compose or read</span></span>|