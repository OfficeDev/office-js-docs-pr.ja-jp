 

# <a name="office"></a><span data-ttu-id="5829f-101">Office</span><span class="sxs-lookup"><span data-stu-id="5829f-101">Office</span></span>

<span data-ttu-id="5829f-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共有 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5829f-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="5829f-104">要件</span><span class="sxs-lookup"><span data-stu-id="5829f-104">Requirements</span></span>

|<span data-ttu-id="5829f-105">要件</span><span class="sxs-lookup"><span data-stu-id="5829f-105">Requirement</span></span>| <span data-ttu-id="5829f-106">値</span><span class="sxs-lookup"><span data-stu-id="5829f-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="5829f-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5829f-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5829f-108">1.0</span><span class="sxs-lookup"><span data-stu-id="5829f-108">1.0</span></span>|
|[<span data-ttu-id="5829f-109">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5829f-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5829f-110">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="5829f-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="5829f-111">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="5829f-111">Members and methods</span></span>

| <span data-ttu-id="5829f-112">メンバー</span><span class="sxs-lookup"><span data-stu-id="5829f-112">Member</span></span> | <span data-ttu-id="5829f-113">種類</span><span class="sxs-lookup"><span data-stu-id="5829f-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="5829f-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="5829f-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="5829f-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="5829f-115">Member</span></span> |
| [<span data-ttu-id="5829f-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="5829f-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="5829f-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="5829f-117">Member</span></span> |
| [<span data-ttu-id="5829f-118">EventType</span><span class="sxs-lookup"><span data-stu-id="5829f-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="5829f-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="5829f-119">Member</span></span> |
| [<span data-ttu-id="5829f-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="5829f-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="5829f-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="5829f-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="5829f-122">名前空間</span><span class="sxs-lookup"><span data-stu-id="5829f-122">Namespaces</span></span>

<span data-ttu-id="5829f-123">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="5829f-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="5829f-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="5829f-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="5829f-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="5829f-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="5829f-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="5829f-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="5829f-127">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="5829f-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="5829f-128">型:</span><span class="sxs-lookup"><span data-stu-id="5829f-128">Type:</span></span>

*   <span data-ttu-id="5829f-129">String</span><span class="sxs-lookup"><span data-stu-id="5829f-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5829f-130">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="5829f-130">Properties:</span></span>

|<span data-ttu-id="5829f-131">名前</span><span class="sxs-lookup"><span data-stu-id="5829f-131">Name</span></span>| <span data-ttu-id="5829f-132">型</span><span class="sxs-lookup"><span data-stu-id="5829f-132">Type</span></span>| <span data-ttu-id="5829f-133">説明</span><span class="sxs-lookup"><span data-stu-id="5829f-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="5829f-134">String</span><span class="sxs-lookup"><span data-stu-id="5829f-134">String</span></span>|<span data-ttu-id="5829f-135">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="5829f-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="5829f-136">String</span><span class="sxs-lookup"><span data-stu-id="5829f-136">String</span></span>|<span data-ttu-id="5829f-137">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="5829f-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5829f-138">要件</span><span class="sxs-lookup"><span data-stu-id="5829f-138">Requirements</span></span>

|<span data-ttu-id="5829f-139">要件</span><span class="sxs-lookup"><span data-stu-id="5829f-139">Requirement</span></span>| <span data-ttu-id="5829f-140">値</span><span class="sxs-lookup"><span data-stu-id="5829f-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="5829f-141">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5829f-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5829f-142">1.0</span><span class="sxs-lookup"><span data-stu-id="5829f-142">1.0</span></span>|
|[<span data-ttu-id="5829f-143">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5829f-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5829f-144">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="5829f-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="5829f-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="5829f-145">CoercionType :String</span></span>

<span data-ttu-id="5829f-146">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="5829f-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5829f-147">型:</span><span class="sxs-lookup"><span data-stu-id="5829f-147">Type:</span></span>

*   <span data-ttu-id="5829f-148">String</span><span class="sxs-lookup"><span data-stu-id="5829f-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5829f-149">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="5829f-149">Properties:</span></span>

|<span data-ttu-id="5829f-150">名前</span><span class="sxs-lookup"><span data-stu-id="5829f-150">Name</span></span>| <span data-ttu-id="5829f-151">型</span><span class="sxs-lookup"><span data-stu-id="5829f-151">Type</span></span>| <span data-ttu-id="5829f-152">説明</span><span class="sxs-lookup"><span data-stu-id="5829f-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="5829f-153">String</span><span class="sxs-lookup"><span data-stu-id="5829f-153">String</span></span>|<span data-ttu-id="5829f-154">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="5829f-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="5829f-155">String</span><span class="sxs-lookup"><span data-stu-id="5829f-155">String</span></span>|<span data-ttu-id="5829f-156">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="5829f-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5829f-157">要件</span><span class="sxs-lookup"><span data-stu-id="5829f-157">Requirements</span></span>

|<span data-ttu-id="5829f-158">要件</span><span class="sxs-lookup"><span data-stu-id="5829f-158">Requirement</span></span>| <span data-ttu-id="5829f-159">値</span><span class="sxs-lookup"><span data-stu-id="5829f-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="5829f-160">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5829f-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5829f-161">1.0</span><span class="sxs-lookup"><span data-stu-id="5829f-161">1.0</span></span>|
|[<span data-ttu-id="5829f-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5829f-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5829f-163">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="5829f-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="5829f-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="5829f-164">EventType :String</span></span>

<span data-ttu-id="5829f-165">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="5829f-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="5829f-166">型:</span><span class="sxs-lookup"><span data-stu-id="5829f-166">Type:</span></span>

*   <span data-ttu-id="5829f-167">String</span><span class="sxs-lookup"><span data-stu-id="5829f-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5829f-168">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="5829f-168">Properties:</span></span>

| <span data-ttu-id="5829f-169">名前</span><span class="sxs-lookup"><span data-stu-id="5829f-169">Name</span></span> | <span data-ttu-id="5829f-170">型</span><span class="sxs-lookup"><span data-stu-id="5829f-170">Type</span></span> | <span data-ttu-id="5829f-171">説明</span><span class="sxs-lookup"><span data-stu-id="5829f-171">Description</span></span> | <span data-ttu-id="5829f-172">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="5829f-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="5829f-173">文字列</span><span class="sxs-lookup"><span data-stu-id="5829f-173">String</span></span> | <span data-ttu-id="5829f-174">選択した予定または一連の予定の日付または時刻が変更された。</span><span class="sxs-lookup"><span data-stu-id="5829f-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="5829f-175">1.7</span><span class="sxs-lookup"><span data-stu-id="5829f-175">-17</span></span> |
|`ItemChanged`| <span data-ttu-id="5829f-176">文字列</span><span class="sxs-lookup"><span data-stu-id="5829f-176">String</span></span> | <span data-ttu-id="5829f-177">作業ウィンドウがピン留めされている間、別の Outlook アイテムが選択されて表示されている。</span><span class="sxs-lookup"><span data-stu-id="5829f-177">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="5829f-178">1.5</span><span class="sxs-lookup"><span data-stu-id="5829f-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="5829f-179">文字列</span><span class="sxs-lookup"><span data-stu-id="5829f-179">String</span></span> | <span data-ttu-id="5829f-180">選択したアイテムまたは予定の場所の受信者リストが変更された。</span><span class="sxs-lookup"><span data-stu-id="5829f-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="5829f-181">1.7</span><span class="sxs-lookup"><span data-stu-id="5829f-181">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="5829f-182">文字列</span><span class="sxs-lookup"><span data-stu-id="5829f-182">String</span></span> | <span data-ttu-id="5829f-183">選択した一連の予定の定期的なパターンが変更された。</span><span class="sxs-lookup"><span data-stu-id="5829f-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="5829f-184">1.7</span><span class="sxs-lookup"><span data-stu-id="5829f-184">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5829f-185">要件</span><span class="sxs-lookup"><span data-stu-id="5829f-185">Requirements</span></span>

|<span data-ttu-id="5829f-186">要件</span><span class="sxs-lookup"><span data-stu-id="5829f-186">Requirement</span></span>| <span data-ttu-id="5829f-187">値</span><span class="sxs-lookup"><span data-stu-id="5829f-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="5829f-188">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5829f-188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5829f-189">1.5</span><span class="sxs-lookup"><span data-stu-id="5829f-189">1.5</span></span> |
|[<span data-ttu-id="5829f-190">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5829f-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5829f-191">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="5829f-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="5829f-192">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="5829f-192">SourceProperty :String</span></span>

<span data-ttu-id="5829f-193">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="5829f-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5829f-194">型:</span><span class="sxs-lookup"><span data-stu-id="5829f-194">Type:</span></span>

*   <span data-ttu-id="5829f-195">String</span><span class="sxs-lookup"><span data-stu-id="5829f-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5829f-196">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="5829f-196">Properties:</span></span>

|<span data-ttu-id="5829f-197">名前</span><span class="sxs-lookup"><span data-stu-id="5829f-197">Name</span></span>| <span data-ttu-id="5829f-198">型</span><span class="sxs-lookup"><span data-stu-id="5829f-198">Type</span></span>| <span data-ttu-id="5829f-199">説明</span><span class="sxs-lookup"><span data-stu-id="5829f-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="5829f-200">String</span><span class="sxs-lookup"><span data-stu-id="5829f-200">String</span></span>|<span data-ttu-id="5829f-201">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="5829f-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="5829f-202">String</span><span class="sxs-lookup"><span data-stu-id="5829f-202">String</span></span>|<span data-ttu-id="5829f-203">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="5829f-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5829f-204">要件</span><span class="sxs-lookup"><span data-stu-id="5829f-204">Requirements</span></span>

|<span data-ttu-id="5829f-205">要件</span><span class="sxs-lookup"><span data-stu-id="5829f-205">Requirement</span></span>| <span data-ttu-id="5829f-206">値</span><span class="sxs-lookup"><span data-stu-id="5829f-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="5829f-207">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5829f-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5829f-208">1.0</span><span class="sxs-lookup"><span data-stu-id="5829f-208">1.0</span></span>|
|[<span data-ttu-id="5829f-209">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5829f-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5829f-210">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="5829f-210">Compose or read</span></span>|