# <a name="office"></a><span data-ttu-id="989c2-101">Office</span><span class="sxs-lookup"><span data-stu-id="989c2-101">Office</span></span>

<span data-ttu-id="989c2-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共有 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="989c2-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="989c2-104">要件</span><span class="sxs-lookup"><span data-stu-id="989c2-104">Requirements</span></span>

|<span data-ttu-id="989c2-105">要件</span><span class="sxs-lookup"><span data-stu-id="989c2-105">Requirement</span></span>| <span data-ttu-id="989c2-106">値</span><span class="sxs-lookup"><span data-stu-id="989c2-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="989c2-107">メールボックス要件の最小バージョン</span><span class="sxs-lookup"><span data-stu-id="989c2-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="989c2-108">1.0</span><span class="sxs-lookup"><span data-stu-id="989c2-108">1.0</span></span>|
|[<span data-ttu-id="989c2-109">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="989c2-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="989c2-110">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="989c2-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="989c2-111">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="989c2-111">Members and methods</span></span>

| <span data-ttu-id="989c2-112">メンバー</span><span class="sxs-lookup"><span data-stu-id="989c2-112">Member</span></span> | <span data-ttu-id="989c2-113">種類</span><span class="sxs-lookup"><span data-stu-id="989c2-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="989c2-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="989c2-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="989c2-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="989c2-115">Member</span></span> |
| [<span data-ttu-id="989c2-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="989c2-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="989c2-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="989c2-117">Member</span></span> |
| [<span data-ttu-id="989c2-118">EventType</span><span class="sxs-lookup"><span data-stu-id="989c2-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="989c2-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="989c2-119">Member</span></span> |
| [<span data-ttu-id="989c2-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="989c2-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="989c2-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="989c2-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="989c2-122">名前空間</span><span class="sxs-lookup"><span data-stu-id="989c2-122">Namespaces</span></span>

<span data-ttu-id="989c2-123">[context](office.context.md):Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="989c2-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="989c2-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype):ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="989c2-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="989c2-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="989c2-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="989c2-126">AsyncResultStatus: 文字列</span><span class="sxs-lookup"><span data-stu-id="989c2-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="989c2-127">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="989c2-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="989c2-128">種類:</span><span class="sxs-lookup"><span data-stu-id="989c2-128">Type:</span></span>

*   <span data-ttu-id="989c2-129">文字列</span><span class="sxs-lookup"><span data-stu-id="989c2-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="989c2-130">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="989c2-130">Properties:</span></span>

|<span data-ttu-id="989c2-131">名前</span><span class="sxs-lookup"><span data-stu-id="989c2-131">Name</span></span>| <span data-ttu-id="989c2-132">種類</span><span class="sxs-lookup"><span data-stu-id="989c2-132">Type</span></span>| <span data-ttu-id="989c2-133">説明</span><span class="sxs-lookup"><span data-stu-id="989c2-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="989c2-134">文字列</span><span class="sxs-lookup"><span data-stu-id="989c2-134">String</span></span>|<span data-ttu-id="989c2-135">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="989c2-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="989c2-136">文字列</span><span class="sxs-lookup"><span data-stu-id="989c2-136">String</span></span>|<span data-ttu-id="989c2-137">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="989c2-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="989c2-138">要件</span><span class="sxs-lookup"><span data-stu-id="989c2-138">Requirements</span></span>

|<span data-ttu-id="989c2-139">要件</span><span class="sxs-lookup"><span data-stu-id="989c2-139">Requirement</span></span>| <span data-ttu-id="989c2-140">値</span><span class="sxs-lookup"><span data-stu-id="989c2-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="989c2-141">メールボックス要件の最小バージョン</span><span class="sxs-lookup"><span data-stu-id="989c2-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="989c2-142">1.0</span><span class="sxs-lookup"><span data-stu-id="989c2-142">1.0</span></span>|
|[<span data-ttu-id="989c2-143">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="989c2-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="989c2-144">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="989c2-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="989c2-145">CoercionType: 文字列</span><span class="sxs-lookup"><span data-stu-id="989c2-145">CoercionType :String</span></span>

<span data-ttu-id="989c2-146">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="989c2-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="989c2-147">種類:</span><span class="sxs-lookup"><span data-stu-id="989c2-147">Type:</span></span>

*   <span data-ttu-id="989c2-148">文字列</span><span class="sxs-lookup"><span data-stu-id="989c2-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="989c2-149">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="989c2-149">Properties:</span></span>

|<span data-ttu-id="989c2-150">名前</span><span class="sxs-lookup"><span data-stu-id="989c2-150">Name</span></span>| <span data-ttu-id="989c2-151">種類</span><span class="sxs-lookup"><span data-stu-id="989c2-151">Type</span></span>| <span data-ttu-id="989c2-152">説明</span><span class="sxs-lookup"><span data-stu-id="989c2-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="989c2-153">文字列</span><span class="sxs-lookup"><span data-stu-id="989c2-153">String</span></span>|<span data-ttu-id="989c2-154">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="989c2-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="989c2-155">文字列</span><span class="sxs-lookup"><span data-stu-id="989c2-155">String</span></span>|<span data-ttu-id="989c2-156">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="989c2-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="989c2-157">要件</span><span class="sxs-lookup"><span data-stu-id="989c2-157">Requirements</span></span>

|<span data-ttu-id="989c2-158">要件</span><span class="sxs-lookup"><span data-stu-id="989c2-158">Requirement</span></span>| <span data-ttu-id="989c2-159">値</span><span class="sxs-lookup"><span data-stu-id="989c2-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="989c2-160">メールボックス要件の最小バージョン</span><span class="sxs-lookup"><span data-stu-id="989c2-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="989c2-161">1.0</span><span class="sxs-lookup"><span data-stu-id="989c2-161">1.0</span></span>|
|[<span data-ttu-id="989c2-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="989c2-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="989c2-163">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="989c2-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="989c2-164">イベントの種類: 文字列</span><span class="sxs-lookup"><span data-stu-id="989c2-164">EventType :String</span></span>

<span data-ttu-id="989c2-165">イベント ハンドラに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="989c2-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="989c2-166">種類:</span><span class="sxs-lookup"><span data-stu-id="989c2-166">Type:</span></span>

*   <span data-ttu-id="989c2-167">文字列</span><span class="sxs-lookup"><span data-stu-id="989c2-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="989c2-168">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="989c2-168">Properties:</span></span>

| <span data-ttu-id="989c2-169">名前</span><span class="sxs-lookup"><span data-stu-id="989c2-169">Name</span></span> | <span data-ttu-id="989c2-170">種類</span><span class="sxs-lookup"><span data-stu-id="989c2-170">Type</span></span> | <span data-ttu-id="989c2-171">説明</span><span class="sxs-lookup"><span data-stu-id="989c2-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="989c2-172">文字列</span><span class="sxs-lookup"><span data-stu-id="989c2-172">String</span></span> | <span data-ttu-id="989c2-173">選択したアイテムが変更されました。</span><span class="sxs-lookup"><span data-stu-id="989c2-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="989c2-174">要件</span><span class="sxs-lookup"><span data-stu-id="989c2-174">Requirements</span></span>

|<span data-ttu-id="989c2-175">要件</span><span class="sxs-lookup"><span data-stu-id="989c2-175">Requirement</span></span>| <span data-ttu-id="989c2-176">値</span><span class="sxs-lookup"><span data-stu-id="989c2-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="989c2-177">メールボックス要件の最小バージョン</span><span class="sxs-lookup"><span data-stu-id="989c2-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="989c2-178">1.5</span><span class="sxs-lookup"><span data-stu-id="989c2-178">1.5</span></span> |
|[<span data-ttu-id="989c2-179">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="989c2-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="989c2-180">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="989c2-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="989c2-181">SourceProperty: 文字列</span><span class="sxs-lookup"><span data-stu-id="989c2-181">SourceProperty :String</span></span>

<span data-ttu-id="989c2-182">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="989c2-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="989c2-183">種類:</span><span class="sxs-lookup"><span data-stu-id="989c2-183">Type:</span></span>

*   <span data-ttu-id="989c2-184">文字列</span><span class="sxs-lookup"><span data-stu-id="989c2-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="989c2-185">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="989c2-185">Properties:</span></span>

|<span data-ttu-id="989c2-186">名前</span><span class="sxs-lookup"><span data-stu-id="989c2-186">Name</span></span>| <span data-ttu-id="989c2-187">種類</span><span class="sxs-lookup"><span data-stu-id="989c2-187">Type</span></span>| <span data-ttu-id="989c2-188">説明</span><span class="sxs-lookup"><span data-stu-id="989c2-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="989c2-189">文字列</span><span class="sxs-lookup"><span data-stu-id="989c2-189">String</span></span>|<span data-ttu-id="989c2-190">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="989c2-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="989c2-191">文字列</span><span class="sxs-lookup"><span data-stu-id="989c2-191">String</span></span>|<span data-ttu-id="989c2-192">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="989c2-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="989c2-193">要件</span><span class="sxs-lookup"><span data-stu-id="989c2-193">Requirements</span></span>

|<span data-ttu-id="989c2-194">要件</span><span class="sxs-lookup"><span data-stu-id="989c2-194">Requirement</span></span>| <span data-ttu-id="989c2-195">値</span><span class="sxs-lookup"><span data-stu-id="989c2-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="989c2-196">メールボックス要件の最小バージョン</span><span class="sxs-lookup"><span data-stu-id="989c2-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="989c2-197">1.0</span><span class="sxs-lookup"><span data-stu-id="989c2-197">1.0</span></span>|
|[<span data-ttu-id="989c2-198">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="989c2-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="989c2-199">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="989c2-199">Compose or read</span></span>|