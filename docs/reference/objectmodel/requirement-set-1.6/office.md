 

# <a name="office"></a><span data-ttu-id="df7c3-101">Office</span><span class="sxs-lookup"><span data-stu-id="df7c3-101">Office</span></span>

<span data-ttu-id="df7c3-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共有 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="df7c3-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="df7c3-104">要件</span><span class="sxs-lookup"><span data-stu-id="df7c3-104">Requirements</span></span>

|<span data-ttu-id="df7c3-105">要件</span><span class="sxs-lookup"><span data-stu-id="df7c3-105">Requirement</span></span>| <span data-ttu-id="df7c3-106">値</span><span class="sxs-lookup"><span data-stu-id="df7c3-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="df7c3-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="df7c3-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="df7c3-108">1.0</span><span class="sxs-lookup"><span data-stu-id="df7c3-108">1.0</span></span>|
|[<span data-ttu-id="df7c3-109">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="df7c3-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="df7c3-110">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="df7c3-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="df7c3-111">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="df7c3-111">Members and methods</span></span>

| <span data-ttu-id="df7c3-112">メンバー</span><span class="sxs-lookup"><span data-stu-id="df7c3-112">Member</span></span> | <span data-ttu-id="df7c3-113">種類</span><span class="sxs-lookup"><span data-stu-id="df7c3-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="df7c3-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="df7c3-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="df7c3-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="df7c3-115">Member</span></span> |
| [<span data-ttu-id="df7c3-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="df7c3-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="df7c3-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="df7c3-117">Member</span></span> |
| [<span data-ttu-id="df7c3-118">EventType</span><span class="sxs-lookup"><span data-stu-id="df7c3-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="df7c3-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="df7c3-119">Member</span></span> |
| [<span data-ttu-id="df7c3-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="df7c3-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="df7c3-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="df7c3-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="df7c3-122">名前空間</span><span class="sxs-lookup"><span data-stu-id="df7c3-122">Namespaces</span></span>

<span data-ttu-id="df7c3-123">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="df7c3-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="df7c3-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="df7c3-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="df7c3-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="df7c3-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="df7c3-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="df7c3-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="df7c3-127">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="df7c3-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="df7c3-128">型:</span><span class="sxs-lookup"><span data-stu-id="df7c3-128">Type:</span></span>

*   <span data-ttu-id="df7c3-129">String</span><span class="sxs-lookup"><span data-stu-id="df7c3-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="df7c3-130">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="df7c3-130">Properties:</span></span>

|<span data-ttu-id="df7c3-131">名前</span><span class="sxs-lookup"><span data-stu-id="df7c3-131">Name</span></span>| <span data-ttu-id="df7c3-132">型</span><span class="sxs-lookup"><span data-stu-id="df7c3-132">Type</span></span>| <span data-ttu-id="df7c3-133">説明</span><span class="sxs-lookup"><span data-stu-id="df7c3-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="df7c3-134">String</span><span class="sxs-lookup"><span data-stu-id="df7c3-134">String</span></span>|<span data-ttu-id="df7c3-135">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="df7c3-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="df7c3-136">String</span><span class="sxs-lookup"><span data-stu-id="df7c3-136">String</span></span>|<span data-ttu-id="df7c3-137">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="df7c3-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="df7c3-138">要件</span><span class="sxs-lookup"><span data-stu-id="df7c3-138">Requirements</span></span>

|<span data-ttu-id="df7c3-139">要件</span><span class="sxs-lookup"><span data-stu-id="df7c3-139">Requirement</span></span>| <span data-ttu-id="df7c3-140">値</span><span class="sxs-lookup"><span data-stu-id="df7c3-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="df7c3-141">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="df7c3-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="df7c3-142">1.0</span><span class="sxs-lookup"><span data-stu-id="df7c3-142">1.0</span></span>|
|[<span data-ttu-id="df7c3-143">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="df7c3-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="df7c3-144">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="df7c3-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="df7c3-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="df7c3-145">CoercionType :String</span></span>

<span data-ttu-id="df7c3-146">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="df7c3-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="df7c3-147">型:</span><span class="sxs-lookup"><span data-stu-id="df7c3-147">Type:</span></span>

*   <span data-ttu-id="df7c3-148">String</span><span class="sxs-lookup"><span data-stu-id="df7c3-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="df7c3-149">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="df7c3-149">Properties:</span></span>

|<span data-ttu-id="df7c3-150">名前</span><span class="sxs-lookup"><span data-stu-id="df7c3-150">Name</span></span>| <span data-ttu-id="df7c3-151">型</span><span class="sxs-lookup"><span data-stu-id="df7c3-151">Type</span></span>| <span data-ttu-id="df7c3-152">説明</span><span class="sxs-lookup"><span data-stu-id="df7c3-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="df7c3-153">String</span><span class="sxs-lookup"><span data-stu-id="df7c3-153">String</span></span>|<span data-ttu-id="df7c3-154">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="df7c3-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="df7c3-155">String</span><span class="sxs-lookup"><span data-stu-id="df7c3-155">String</span></span>|<span data-ttu-id="df7c3-156">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="df7c3-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="df7c3-157">要件</span><span class="sxs-lookup"><span data-stu-id="df7c3-157">Requirements</span></span>

|<span data-ttu-id="df7c3-158">要件</span><span class="sxs-lookup"><span data-stu-id="df7c3-158">Requirement</span></span>| <span data-ttu-id="df7c3-159">値</span><span class="sxs-lookup"><span data-stu-id="df7c3-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="df7c3-160">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="df7c3-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="df7c3-161">1.0</span><span class="sxs-lookup"><span data-stu-id="df7c3-161">1.0</span></span>|
|[<span data-ttu-id="df7c3-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="df7c3-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="df7c3-163">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="df7c3-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="df7c3-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="df7c3-164">EventType :String</span></span>

<span data-ttu-id="df7c3-165">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="df7c3-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="df7c3-166">型:</span><span class="sxs-lookup"><span data-stu-id="df7c3-166">Type:</span></span>

*   <span data-ttu-id="df7c3-167">String</span><span class="sxs-lookup"><span data-stu-id="df7c3-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="df7c3-168">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="df7c3-168">Properties:</span></span>

| <span data-ttu-id="df7c3-169">名前</span><span class="sxs-lookup"><span data-stu-id="df7c3-169">Name</span></span> | <span data-ttu-id="df7c3-170">型</span><span class="sxs-lookup"><span data-stu-id="df7c3-170">Type</span></span> | <span data-ttu-id="df7c3-171">説明</span><span class="sxs-lookup"><span data-stu-id="df7c3-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="df7c3-172">文字列</span><span class="sxs-lookup"><span data-stu-id="df7c3-172">String</span></span> | <span data-ttu-id="df7c3-173">作業ウィンドウがピン留めされている間、別の Outlook アイテムが選択されて表示されている。</span><span class="sxs-lookup"><span data-stu-id="df7c3-173">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="df7c3-174">要件</span><span class="sxs-lookup"><span data-stu-id="df7c3-174">Requirements</span></span>

|<span data-ttu-id="df7c3-175">要件</span><span class="sxs-lookup"><span data-stu-id="df7c3-175">Requirement</span></span>| <span data-ttu-id="df7c3-176">値</span><span class="sxs-lookup"><span data-stu-id="df7c3-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="df7c3-177">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="df7c3-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="df7c3-178">1.5</span><span class="sxs-lookup"><span data-stu-id="df7c3-178">1.5</span></span> |
|[<span data-ttu-id="df7c3-179">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="df7c3-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="df7c3-180">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="df7c3-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="df7c3-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="df7c3-181">SourceProperty :String</span></span>

<span data-ttu-id="df7c3-182">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="df7c3-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="df7c3-183">型:</span><span class="sxs-lookup"><span data-stu-id="df7c3-183">Type:</span></span>

*   <span data-ttu-id="df7c3-184">String</span><span class="sxs-lookup"><span data-stu-id="df7c3-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="df7c3-185">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="df7c3-185">Properties:</span></span>

|<span data-ttu-id="df7c3-186">名前</span><span class="sxs-lookup"><span data-stu-id="df7c3-186">Name</span></span>| <span data-ttu-id="df7c3-187">型</span><span class="sxs-lookup"><span data-stu-id="df7c3-187">Type</span></span>| <span data-ttu-id="df7c3-188">説明</span><span class="sxs-lookup"><span data-stu-id="df7c3-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="df7c3-189">String</span><span class="sxs-lookup"><span data-stu-id="df7c3-189">String</span></span>|<span data-ttu-id="df7c3-190">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="df7c3-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="df7c3-191">String</span><span class="sxs-lookup"><span data-stu-id="df7c3-191">String</span></span>|<span data-ttu-id="df7c3-192">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="df7c3-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="df7c3-193">要件</span><span class="sxs-lookup"><span data-stu-id="df7c3-193">Requirements</span></span>

|<span data-ttu-id="df7c3-194">要件</span><span class="sxs-lookup"><span data-stu-id="df7c3-194">Requirement</span></span>| <span data-ttu-id="df7c3-195">値</span><span class="sxs-lookup"><span data-stu-id="df7c3-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="df7c3-196">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="df7c3-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="df7c3-197">1.0</span><span class="sxs-lookup"><span data-stu-id="df7c3-197">1.0</span></span>|
|[<span data-ttu-id="df7c3-198">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="df7c3-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="df7c3-199">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="df7c3-199">Compose or read</span></span>|