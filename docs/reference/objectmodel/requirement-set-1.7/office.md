 

# <a name="office"></a><span data-ttu-id="31ce0-101">Office</span><span class="sxs-lookup"><span data-stu-id="31ce0-101">Office</span></span>

<span data-ttu-id="31ce0-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共有 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="31ce0-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="31ce0-104">要件</span><span class="sxs-lookup"><span data-stu-id="31ce0-104">Requirements</span></span>

|<span data-ttu-id="31ce0-105">要件</span><span class="sxs-lookup"><span data-stu-id="31ce0-105">Requirement</span></span>| <span data-ttu-id="31ce0-106">値</span><span class="sxs-lookup"><span data-stu-id="31ce0-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="31ce0-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="31ce0-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31ce0-108">1.0</span><span class="sxs-lookup"><span data-stu-id="31ce0-108">1.0</span></span>|
|[<span data-ttu-id="31ce0-109">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="31ce0-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31ce0-110">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="31ce0-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="31ce0-111">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="31ce0-111">Members and methods</span></span>

| <span data-ttu-id="31ce0-112">メンバー</span><span class="sxs-lookup"><span data-stu-id="31ce0-112">Member</span></span> | <span data-ttu-id="31ce0-113">型</span><span class="sxs-lookup"><span data-stu-id="31ce0-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="31ce0-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="31ce0-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="31ce0-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="31ce0-115">Member</span></span> |
| [<span data-ttu-id="31ce0-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="31ce0-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="31ce0-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="31ce0-117">Member</span></span> |
| [<span data-ttu-id="31ce0-118">EventType</span><span class="sxs-lookup"><span data-stu-id="31ce0-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="31ce0-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="31ce0-119">Member</span></span> |
| [<span data-ttu-id="31ce0-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="31ce0-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="31ce0-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="31ce0-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="31ce0-122">名前空間</span><span class="sxs-lookup"><span data-stu-id="31ce0-122">Namespaces</span></span>

<span data-ttu-id="31ce0-123">[context](office.context.md):Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="31ce0-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="31ce0-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype):ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="31ce0-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="31ce0-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="31ce0-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="31ce0-126">AsyncResultStatus: 文字列</span><span class="sxs-lookup"><span data-stu-id="31ce0-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="31ce0-127">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="31ce0-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="31ce0-128">型:</span><span class="sxs-lookup"><span data-stu-id="31ce0-128">Type:</span></span>

*   <span data-ttu-id="31ce0-129">文字列</span><span class="sxs-lookup"><span data-stu-id="31ce0-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="31ce0-130">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="31ce0-130">Properties:</span></span>

|<span data-ttu-id="31ce0-131">名前</span><span class="sxs-lookup"><span data-stu-id="31ce0-131">Name</span></span>| <span data-ttu-id="31ce0-132">型</span><span class="sxs-lookup"><span data-stu-id="31ce0-132">Type</span></span>| <span data-ttu-id="31ce0-133">説明</span><span class="sxs-lookup"><span data-stu-id="31ce0-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="31ce0-134">文字列</span><span class="sxs-lookup"><span data-stu-id="31ce0-134">String</span></span>|<span data-ttu-id="31ce0-135">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="31ce0-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="31ce0-136">文字列</span><span class="sxs-lookup"><span data-stu-id="31ce0-136">String</span></span>|<span data-ttu-id="31ce0-137">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="31ce0-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="31ce0-138">要件</span><span class="sxs-lookup"><span data-stu-id="31ce0-138">Requirements</span></span>

|<span data-ttu-id="31ce0-139">要件</span><span class="sxs-lookup"><span data-stu-id="31ce0-139">Requirement</span></span>| <span data-ttu-id="31ce0-140">値</span><span class="sxs-lookup"><span data-stu-id="31ce0-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="31ce0-141">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="31ce0-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31ce0-142">1.0</span><span class="sxs-lookup"><span data-stu-id="31ce0-142">1.0</span></span>|
|[<span data-ttu-id="31ce0-143">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="31ce0-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31ce0-144">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="31ce0-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="31ce0-145">CoercionType: 文字列</span><span class="sxs-lookup"><span data-stu-id="31ce0-145">CoercionType :String</span></span>

<span data-ttu-id="31ce0-146">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="31ce0-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="31ce0-147">型:</span><span class="sxs-lookup"><span data-stu-id="31ce0-147">Type:</span></span>

*   <span data-ttu-id="31ce0-148">文字列</span><span class="sxs-lookup"><span data-stu-id="31ce0-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="31ce0-149">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="31ce0-149">Properties:</span></span>

|<span data-ttu-id="31ce0-150">名前</span><span class="sxs-lookup"><span data-stu-id="31ce0-150">Name</span></span>| <span data-ttu-id="31ce0-151">型</span><span class="sxs-lookup"><span data-stu-id="31ce0-151">Type</span></span>| <span data-ttu-id="31ce0-152">説明</span><span class="sxs-lookup"><span data-stu-id="31ce0-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="31ce0-153">文字列</span><span class="sxs-lookup"><span data-stu-id="31ce0-153">String</span></span>|<span data-ttu-id="31ce0-154">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="31ce0-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="31ce0-155">文字列</span><span class="sxs-lookup"><span data-stu-id="31ce0-155">String</span></span>|<span data-ttu-id="31ce0-156">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="31ce0-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="31ce0-157">要件</span><span class="sxs-lookup"><span data-stu-id="31ce0-157">Requirements</span></span>

|<span data-ttu-id="31ce0-158">要件</span><span class="sxs-lookup"><span data-stu-id="31ce0-158">Requirement</span></span>| <span data-ttu-id="31ce0-159">値</span><span class="sxs-lookup"><span data-stu-id="31ce0-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="31ce0-160">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="31ce0-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31ce0-161">1.0</span><span class="sxs-lookup"><span data-stu-id="31ce0-161">1.0</span></span>|
|[<span data-ttu-id="31ce0-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="31ce0-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31ce0-163">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="31ce0-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="31ce0-164">イベントの種類: 文字列</span><span class="sxs-lookup"><span data-stu-id="31ce0-164">EventType :String</span></span>

<span data-ttu-id="31ce0-165">イベント ハンドラに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="31ce0-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="31ce0-166">型:</span><span class="sxs-lookup"><span data-stu-id="31ce0-166">Type:</span></span>

*   <span data-ttu-id="31ce0-167">文字列</span><span class="sxs-lookup"><span data-stu-id="31ce0-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="31ce0-168">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="31ce0-168">Properties:</span></span>

| <span data-ttu-id="31ce0-169">名前</span><span class="sxs-lookup"><span data-stu-id="31ce0-169">Name</span></span> | <span data-ttu-id="31ce0-170">型</span><span class="sxs-lookup"><span data-stu-id="31ce0-170">Type</span></span> | <span data-ttu-id="31ce0-171">説明</span><span class="sxs-lookup"><span data-stu-id="31ce0-171">Description</span></span> | <span data-ttu-id="31ce0-172">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="31ce0-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="31ce0-173">文字列</span><span class="sxs-lookup"><span data-stu-id="31ce0-173">String</span></span> | <span data-ttu-id="31ce0-174">選択した予定または系列の日付または時間が変更されました。</span><span class="sxs-lookup"><span data-stu-id="31ce0-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="31ce0-175">1.7</span><span class="sxs-lookup"><span data-stu-id="31ce0-175">-17</span></span> |
|`ItemChanged`| <span data-ttu-id="31ce0-176">文字列</span><span class="sxs-lookup"><span data-stu-id="31ce0-176">String</span></span> | <span data-ttu-id="31ce0-177">選択したアイテムが変更されました。</span><span class="sxs-lookup"><span data-stu-id="31ce0-177">The selected item has changed.</span></span> | <span data-ttu-id="31ce0-178">1.5</span><span class="sxs-lookup"><span data-stu-id="31ce0-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="31ce0-179">文字列</span><span class="sxs-lookup"><span data-stu-id="31ce0-179">String</span></span> | <span data-ttu-id="31ce0-180">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="31ce0-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="31ce0-181">1.7</span><span class="sxs-lookup"><span data-stu-id="31ce0-181">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="31ce0-182">文字列</span><span class="sxs-lookup"><span data-stu-id="31ce0-182">String</span></span> | <span data-ttu-id="31ce0-183">選択した系列の定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="31ce0-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="31ce0-184">1.7</span><span class="sxs-lookup"><span data-stu-id="31ce0-184">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="31ce0-185">要件</span><span class="sxs-lookup"><span data-stu-id="31ce0-185">Requirements</span></span>

|<span data-ttu-id="31ce0-186">要件</span><span class="sxs-lookup"><span data-stu-id="31ce0-186">Requirement</span></span>| <span data-ttu-id="31ce0-187">値</span><span class="sxs-lookup"><span data-stu-id="31ce0-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="31ce0-188">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="31ce0-188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31ce0-189">1.5</span><span class="sxs-lookup"><span data-stu-id="31ce0-189">1.5</span></span> |
|[<span data-ttu-id="31ce0-190">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="31ce0-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31ce0-191">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="31ce0-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="31ce0-192">SourceProperty: 文字列</span><span class="sxs-lookup"><span data-stu-id="31ce0-192">SourceProperty :String</span></span>

<span data-ttu-id="31ce0-193">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="31ce0-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="31ce0-194">型:</span><span class="sxs-lookup"><span data-stu-id="31ce0-194">Type:</span></span>

*   <span data-ttu-id="31ce0-195">文字列</span><span class="sxs-lookup"><span data-stu-id="31ce0-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="31ce0-196">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="31ce0-196">Properties:</span></span>

|<span data-ttu-id="31ce0-197">名前</span><span class="sxs-lookup"><span data-stu-id="31ce0-197">Name</span></span>| <span data-ttu-id="31ce0-198">型</span><span class="sxs-lookup"><span data-stu-id="31ce0-198">Type</span></span>| <span data-ttu-id="31ce0-199">説明</span><span class="sxs-lookup"><span data-stu-id="31ce0-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="31ce0-200">文字列</span><span class="sxs-lookup"><span data-stu-id="31ce0-200">String</span></span>|<span data-ttu-id="31ce0-201">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="31ce0-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="31ce0-202">文字列</span><span class="sxs-lookup"><span data-stu-id="31ce0-202">String</span></span>|<span data-ttu-id="31ce0-203">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="31ce0-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="31ce0-204">要件</span><span class="sxs-lookup"><span data-stu-id="31ce0-204">Requirements</span></span>

|<span data-ttu-id="31ce0-205">要件</span><span class="sxs-lookup"><span data-stu-id="31ce0-205">Requirement</span></span>| <span data-ttu-id="31ce0-206">値</span><span class="sxs-lookup"><span data-stu-id="31ce0-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="31ce0-207">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="31ce0-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="31ce0-208">1.0</span><span class="sxs-lookup"><span data-stu-id="31ce0-208">1.0</span></span>|
|[<span data-ttu-id="31ce0-209">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="31ce0-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="31ce0-210">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="31ce0-210">Compose or read</span></span>|