# <a name="office"></a><span data-ttu-id="17539-101">Office</span><span class="sxs-lookup"><span data-stu-id="17539-101">Office</span></span>

<span data-ttu-id="17539-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共有 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="17539-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="17539-104">要件</span><span class="sxs-lookup"><span data-stu-id="17539-104">Requirements</span></span>

|<span data-ttu-id="17539-105">要件</span><span class="sxs-lookup"><span data-stu-id="17539-105">Requirement</span></span>| <span data-ttu-id="17539-106">値</span><span class="sxs-lookup"><span data-stu-id="17539-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="17539-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="17539-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="17539-108">1.0</span><span class="sxs-lookup"><span data-stu-id="17539-108">1.0</span></span>|
|[<span data-ttu-id="17539-109">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="17539-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="17539-110">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="17539-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="17539-111">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="17539-111">Members and methods</span></span>

| <span data-ttu-id="17539-112">メンバー</span><span class="sxs-lookup"><span data-stu-id="17539-112">Member</span></span> | <span data-ttu-id="17539-113">種類</span><span class="sxs-lookup"><span data-stu-id="17539-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="17539-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="17539-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="17539-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="17539-115">Member</span></span> |
| [<span data-ttu-id="17539-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="17539-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="17539-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="17539-117">Member</span></span> |
| [<span data-ttu-id="17539-118">EventType</span><span class="sxs-lookup"><span data-stu-id="17539-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="17539-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="17539-119">Member</span></span> |
| [<span data-ttu-id="17539-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="17539-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="17539-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="17539-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="17539-122">名前空間</span><span class="sxs-lookup"><span data-stu-id="17539-122">Namespaces</span></span>

<span data-ttu-id="17539-123">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="17539-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="17539-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="17539-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="17539-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="17539-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="17539-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="17539-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="17539-127">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="17539-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="17539-128">型:</span><span class="sxs-lookup"><span data-stu-id="17539-128">Type:</span></span>

*   <span data-ttu-id="17539-129">String</span><span class="sxs-lookup"><span data-stu-id="17539-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="17539-130">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="17539-130">Properties:</span></span>

|<span data-ttu-id="17539-131">名前</span><span class="sxs-lookup"><span data-stu-id="17539-131">Name</span></span>| <span data-ttu-id="17539-132">型</span><span class="sxs-lookup"><span data-stu-id="17539-132">Type</span></span>| <span data-ttu-id="17539-133">説明</span><span class="sxs-lookup"><span data-stu-id="17539-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="17539-134">String</span><span class="sxs-lookup"><span data-stu-id="17539-134">String</span></span>|<span data-ttu-id="17539-135">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="17539-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="17539-136">String</span><span class="sxs-lookup"><span data-stu-id="17539-136">String</span></span>|<span data-ttu-id="17539-137">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="17539-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="17539-138">要件</span><span class="sxs-lookup"><span data-stu-id="17539-138">Requirements</span></span>

|<span data-ttu-id="17539-139">要件</span><span class="sxs-lookup"><span data-stu-id="17539-139">Requirement</span></span>| <span data-ttu-id="17539-140">値</span><span class="sxs-lookup"><span data-stu-id="17539-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="17539-141">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="17539-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="17539-142">1.0</span><span class="sxs-lookup"><span data-stu-id="17539-142">1.0</span></span>|
|[<span data-ttu-id="17539-143">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="17539-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="17539-144">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="17539-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="17539-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="17539-145">CoercionType :String</span></span>

<span data-ttu-id="17539-146">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="17539-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="17539-147">型:</span><span class="sxs-lookup"><span data-stu-id="17539-147">Type:</span></span>

*   <span data-ttu-id="17539-148">String</span><span class="sxs-lookup"><span data-stu-id="17539-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="17539-149">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="17539-149">Properties:</span></span>

|<span data-ttu-id="17539-150">名前</span><span class="sxs-lookup"><span data-stu-id="17539-150">Name</span></span>| <span data-ttu-id="17539-151">型</span><span class="sxs-lookup"><span data-stu-id="17539-151">Type</span></span>| <span data-ttu-id="17539-152">説明</span><span class="sxs-lookup"><span data-stu-id="17539-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="17539-153">String</span><span class="sxs-lookup"><span data-stu-id="17539-153">String</span></span>|<span data-ttu-id="17539-154">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="17539-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="17539-155">String</span><span class="sxs-lookup"><span data-stu-id="17539-155">String</span></span>|<span data-ttu-id="17539-156">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="17539-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="17539-157">要件</span><span class="sxs-lookup"><span data-stu-id="17539-157">Requirements</span></span>

|<span data-ttu-id="17539-158">要件</span><span class="sxs-lookup"><span data-stu-id="17539-158">Requirement</span></span>| <span data-ttu-id="17539-159">値</span><span class="sxs-lookup"><span data-stu-id="17539-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="17539-160">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="17539-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="17539-161">1.0</span><span class="sxs-lookup"><span data-stu-id="17539-161">1.0</span></span>|
|[<span data-ttu-id="17539-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="17539-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="17539-163">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="17539-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="17539-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="17539-164">EventType :String</span></span>

<span data-ttu-id="17539-165">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="17539-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="17539-166">型:</span><span class="sxs-lookup"><span data-stu-id="17539-166">Type:</span></span>

*   <span data-ttu-id="17539-167">String</span><span class="sxs-lookup"><span data-stu-id="17539-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="17539-168">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="17539-168">Properties:</span></span>

| <span data-ttu-id="17539-169">名前</span><span class="sxs-lookup"><span data-stu-id="17539-169">Name</span></span> | <span data-ttu-id="17539-170">型</span><span class="sxs-lookup"><span data-stu-id="17539-170">Type</span></span> | <span data-ttu-id="17539-171">説明</span><span class="sxs-lookup"><span data-stu-id="17539-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="17539-172">文字列</span><span class="sxs-lookup"><span data-stu-id="17539-172">String</span></span> | <span data-ttu-id="17539-173">作業ウィンドウがピン留めされている間、別の Outlook アイテムが選択されて表示されている。</span><span class="sxs-lookup"><span data-stu-id="17539-173">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="17539-174">要件</span><span class="sxs-lookup"><span data-stu-id="17539-174">Requirements</span></span>

|<span data-ttu-id="17539-175">要件</span><span class="sxs-lookup"><span data-stu-id="17539-175">Requirement</span></span>| <span data-ttu-id="17539-176">値</span><span class="sxs-lookup"><span data-stu-id="17539-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="17539-177">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="17539-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="17539-178">1.5</span><span class="sxs-lookup"><span data-stu-id="17539-178">1.5</span></span> |
|[<span data-ttu-id="17539-179">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="17539-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="17539-180">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="17539-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="17539-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="17539-181">SourceProperty :String</span></span>

<span data-ttu-id="17539-182">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="17539-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="17539-183">型:</span><span class="sxs-lookup"><span data-stu-id="17539-183">Type:</span></span>

*   <span data-ttu-id="17539-184">String</span><span class="sxs-lookup"><span data-stu-id="17539-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="17539-185">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="17539-185">Properties:</span></span>

|<span data-ttu-id="17539-186">名前</span><span class="sxs-lookup"><span data-stu-id="17539-186">Name</span></span>| <span data-ttu-id="17539-187">型</span><span class="sxs-lookup"><span data-stu-id="17539-187">Type</span></span>| <span data-ttu-id="17539-188">説明</span><span class="sxs-lookup"><span data-stu-id="17539-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="17539-189">String</span><span class="sxs-lookup"><span data-stu-id="17539-189">String</span></span>|<span data-ttu-id="17539-190">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="17539-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="17539-191">String</span><span class="sxs-lookup"><span data-stu-id="17539-191">String</span></span>|<span data-ttu-id="17539-192">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="17539-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="17539-193">要件</span><span class="sxs-lookup"><span data-stu-id="17539-193">Requirements</span></span>

|<span data-ttu-id="17539-194">要件</span><span class="sxs-lookup"><span data-stu-id="17539-194">Requirement</span></span>| <span data-ttu-id="17539-195">値</span><span class="sxs-lookup"><span data-stu-id="17539-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="17539-196">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="17539-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="17539-197">1.0</span><span class="sxs-lookup"><span data-stu-id="17539-197">1.0</span></span>|
|[<span data-ttu-id="17539-198">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="17539-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="17539-199">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="17539-199">Compose or read</span></span>|