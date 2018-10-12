 

# <a name="office"></a><span data-ttu-id="c94cd-101">Office</span><span class="sxs-lookup"><span data-stu-id="c94cd-101">Office</span></span>

<span data-ttu-id="c94cd-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共有 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c94cd-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c94cd-104">要件</span><span class="sxs-lookup"><span data-stu-id="c94cd-104">Requirements</span></span>

|<span data-ttu-id="c94cd-105">要件</span><span class="sxs-lookup"><span data-stu-id="c94cd-105">Requirement</span></span>| <span data-ttu-id="c94cd-106">値</span><span class="sxs-lookup"><span data-stu-id="c94cd-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="c94cd-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c94cd-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c94cd-108">1.0</span><span class="sxs-lookup"><span data-stu-id="c94cd-108">1.0</span></span>|
|[<span data-ttu-id="c94cd-109">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c94cd-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c94cd-110">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c94cd-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c94cd-111">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="c94cd-111">Members and methods</span></span>

| <span data-ttu-id="c94cd-112">メンバー</span><span class="sxs-lookup"><span data-stu-id="c94cd-112">Member</span></span> | <span data-ttu-id="c94cd-113">型</span><span class="sxs-lookup"><span data-stu-id="c94cd-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c94cd-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="c94cd-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="c94cd-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="c94cd-115">Member</span></span> |
| [<span data-ttu-id="c94cd-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="c94cd-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="c94cd-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="c94cd-117">Member</span></span> |
| [<span data-ttu-id="c94cd-118">EventType</span><span class="sxs-lookup"><span data-stu-id="c94cd-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="c94cd-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="c94cd-119">Member</span></span> |
| [<span data-ttu-id="c94cd-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="c94cd-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="c94cd-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="c94cd-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="c94cd-122">名前空間</span><span class="sxs-lookup"><span data-stu-id="c94cd-122">Namespaces</span></span>

<span data-ttu-id="c94cd-123">[context](office.context.md):Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c94cd-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="c94cd-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype):ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="c94cd-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="c94cd-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="c94cd-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="c94cd-126">AsyncResultStatus: 文字列</span><span class="sxs-lookup"><span data-stu-id="c94cd-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="c94cd-127">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="c94cd-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="c94cd-128">型:</span><span class="sxs-lookup"><span data-stu-id="c94cd-128">Type:</span></span>

*   <span data-ttu-id="c94cd-129">文字列</span><span class="sxs-lookup"><span data-stu-id="c94cd-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c94cd-130">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="c94cd-130">Properties:</span></span>

|<span data-ttu-id="c94cd-131">名前</span><span class="sxs-lookup"><span data-stu-id="c94cd-131">Name</span></span>| <span data-ttu-id="c94cd-132">型</span><span class="sxs-lookup"><span data-stu-id="c94cd-132">Type</span></span>| <span data-ttu-id="c94cd-133">説明</span><span class="sxs-lookup"><span data-stu-id="c94cd-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="c94cd-134">文字列</span><span class="sxs-lookup"><span data-stu-id="c94cd-134">String</span></span>|<span data-ttu-id="c94cd-135">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="c94cd-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="c94cd-136">文字列</span><span class="sxs-lookup"><span data-stu-id="c94cd-136">String</span></span>|<span data-ttu-id="c94cd-137">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="c94cd-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c94cd-138">要件</span><span class="sxs-lookup"><span data-stu-id="c94cd-138">Requirements</span></span>

|<span data-ttu-id="c94cd-139">要件</span><span class="sxs-lookup"><span data-stu-id="c94cd-139">Requirement</span></span>| <span data-ttu-id="c94cd-140">値</span><span class="sxs-lookup"><span data-stu-id="c94cd-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="c94cd-141">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c94cd-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c94cd-142">1.0</span><span class="sxs-lookup"><span data-stu-id="c94cd-142">1.0</span></span>|
|[<span data-ttu-id="c94cd-143">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c94cd-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c94cd-144">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c94cd-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="c94cd-145">CoercionType: 文字列</span><span class="sxs-lookup"><span data-stu-id="c94cd-145">CoercionType :String</span></span>

<span data-ttu-id="c94cd-146">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="c94cd-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c94cd-147">型:</span><span class="sxs-lookup"><span data-stu-id="c94cd-147">Type:</span></span>

*   <span data-ttu-id="c94cd-148">文字列</span><span class="sxs-lookup"><span data-stu-id="c94cd-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c94cd-149">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="c94cd-149">Properties:</span></span>

|<span data-ttu-id="c94cd-150">名前</span><span class="sxs-lookup"><span data-stu-id="c94cd-150">Name</span></span>| <span data-ttu-id="c94cd-151">型</span><span class="sxs-lookup"><span data-stu-id="c94cd-151">Type</span></span>| <span data-ttu-id="c94cd-152">説明</span><span class="sxs-lookup"><span data-stu-id="c94cd-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="c94cd-153">文字列</span><span class="sxs-lookup"><span data-stu-id="c94cd-153">String</span></span>|<span data-ttu-id="c94cd-154">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="c94cd-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="c94cd-155">文字列</span><span class="sxs-lookup"><span data-stu-id="c94cd-155">String</span></span>|<span data-ttu-id="c94cd-156">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="c94cd-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c94cd-157">要件</span><span class="sxs-lookup"><span data-stu-id="c94cd-157">Requirements</span></span>

|<span data-ttu-id="c94cd-158">要件</span><span class="sxs-lookup"><span data-stu-id="c94cd-158">Requirement</span></span>| <span data-ttu-id="c94cd-159">値</span><span class="sxs-lookup"><span data-stu-id="c94cd-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="c94cd-160">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c94cd-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c94cd-161">1.0</span><span class="sxs-lookup"><span data-stu-id="c94cd-161">1.0</span></span>|
|[<span data-ttu-id="c94cd-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c94cd-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c94cd-163">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c94cd-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="c94cd-164">イベントの種類: 文字列</span><span class="sxs-lookup"><span data-stu-id="c94cd-164">EventType :String</span></span>

<span data-ttu-id="c94cd-165">イベント ハンドラに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="c94cd-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="c94cd-166">型:</span><span class="sxs-lookup"><span data-stu-id="c94cd-166">Type:</span></span>

*   <span data-ttu-id="c94cd-167">文字列</span><span class="sxs-lookup"><span data-stu-id="c94cd-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c94cd-168">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="c94cd-168">Properties:</span></span>

| <span data-ttu-id="c94cd-169">名前</span><span class="sxs-lookup"><span data-stu-id="c94cd-169">Name</span></span> | <span data-ttu-id="c94cd-170">型</span><span class="sxs-lookup"><span data-stu-id="c94cd-170">Type</span></span> | <span data-ttu-id="c94cd-171">説明</span><span class="sxs-lookup"><span data-stu-id="c94cd-171">Description</span></span> | <span data-ttu-id="c94cd-172">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="c94cd-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="c94cd-173">文字列</span><span class="sxs-lookup"><span data-stu-id="c94cd-173">String</span></span> | <span data-ttu-id="c94cd-174">選択した予定または系列の日付または時間が変更されました。</span><span class="sxs-lookup"><span data-stu-id="c94cd-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="c94cd-175">1.7</span><span class="sxs-lookup"><span data-stu-id="c94cd-175">-17</span></span> |
|`ItemChanged`| <span data-ttu-id="c94cd-176">文字列</span><span class="sxs-lookup"><span data-stu-id="c94cd-176">String</span></span> | <span data-ttu-id="c94cd-177">選択したアイテムが変更されました。</span><span class="sxs-lookup"><span data-stu-id="c94cd-177">The selected item has changed.</span></span> | <span data-ttu-id="c94cd-178">1.5</span><span class="sxs-lookup"><span data-stu-id="c94cd-178">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="c94cd-179">文字列</span><span class="sxs-lookup"><span data-stu-id="c94cd-179">String</span></span> | <span data-ttu-id="c94cd-180">選択したアイテムが変更されました。</span><span class="sxs-lookup"><span data-stu-id="c94cd-180">The selected item has changed.</span></span> | <span data-ttu-id="c94cd-181">プレビュー</span><span class="sxs-lookup"><span data-stu-id="c94cd-181">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="c94cd-182">文字列</span><span class="sxs-lookup"><span data-stu-id="c94cd-182">String</span></span> | <span data-ttu-id="c94cd-183">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="c94cd-183">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="c94cd-184">1.7</span><span class="sxs-lookup"><span data-stu-id="c94cd-184">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="c94cd-185">文字列</span><span class="sxs-lookup"><span data-stu-id="c94cd-185">String</span></span> | <span data-ttu-id="c94cd-186">選択した系列の定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="c94cd-186">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="c94cd-187">1.7</span><span class="sxs-lookup"><span data-stu-id="c94cd-187">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c94cd-188">要件</span><span class="sxs-lookup"><span data-stu-id="c94cd-188">Requirements</span></span>

|<span data-ttu-id="c94cd-189">要件</span><span class="sxs-lookup"><span data-stu-id="c94cd-189">Requirement</span></span>| <span data-ttu-id="c94cd-190">値</span><span class="sxs-lookup"><span data-stu-id="c94cd-190">Value</span></span>|
|---|---|
|[<span data-ttu-id="c94cd-191">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c94cd-191">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c94cd-192">1.5</span><span class="sxs-lookup"><span data-stu-id="c94cd-192">1.5</span></span> |
|[<span data-ttu-id="c94cd-193">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c94cd-193">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c94cd-194">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c94cd-194">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="c94cd-195">SourceProperty: 文字列</span><span class="sxs-lookup"><span data-stu-id="c94cd-195">SourceProperty :String</span></span>

<span data-ttu-id="c94cd-196">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="c94cd-196">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c94cd-197">型:</span><span class="sxs-lookup"><span data-stu-id="c94cd-197">Type:</span></span>

*   <span data-ttu-id="c94cd-198">文字列</span><span class="sxs-lookup"><span data-stu-id="c94cd-198">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c94cd-199">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="c94cd-199">Properties:</span></span>

|<span data-ttu-id="c94cd-200">名前</span><span class="sxs-lookup"><span data-stu-id="c94cd-200">Name</span></span>| <span data-ttu-id="c94cd-201">型</span><span class="sxs-lookup"><span data-stu-id="c94cd-201">Type</span></span>| <span data-ttu-id="c94cd-202">説明</span><span class="sxs-lookup"><span data-stu-id="c94cd-202">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="c94cd-203">文字列</span><span class="sxs-lookup"><span data-stu-id="c94cd-203">String</span></span>|<span data-ttu-id="c94cd-204">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="c94cd-204">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="c94cd-205">文字列</span><span class="sxs-lookup"><span data-stu-id="c94cd-205">String</span></span>|<span data-ttu-id="c94cd-206">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="c94cd-206">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c94cd-207">要件</span><span class="sxs-lookup"><span data-stu-id="c94cd-207">Requirements</span></span>

|<span data-ttu-id="c94cd-208">要件</span><span class="sxs-lookup"><span data-stu-id="c94cd-208">Requirement</span></span>| <span data-ttu-id="c94cd-209">値</span><span class="sxs-lookup"><span data-stu-id="c94cd-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="c94cd-210">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c94cd-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c94cd-211">1.0</span><span class="sxs-lookup"><span data-stu-id="c94cd-211">1.0</span></span>|
|[<span data-ttu-id="c94cd-212">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c94cd-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c94cd-213">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c94cd-213">Compose or read</span></span>|