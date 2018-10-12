 

# <a name="office"></a><span data-ttu-id="9cda4-101">Office</span><span class="sxs-lookup"><span data-stu-id="9cda4-101">Office</span></span>

<span data-ttu-id="9cda4-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共有 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9cda4-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="9cda4-104">要件</span><span class="sxs-lookup"><span data-stu-id="9cda4-104">Requirements</span></span>

|<span data-ttu-id="9cda4-105">要件</span><span class="sxs-lookup"><span data-stu-id="9cda4-105">Requirement</span></span>| <span data-ttu-id="9cda4-106">値</span><span class="sxs-lookup"><span data-stu-id="9cda4-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="9cda4-107">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="9cda4-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9cda4-108">1.0</span><span class="sxs-lookup"><span data-stu-id="9cda4-108">1.0</span></span>|
|[<span data-ttu-id="9cda4-109">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9cda4-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9cda4-110">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="9cda4-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="9cda4-111">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="9cda4-111">Members and methods</span></span>

| <span data-ttu-id="9cda4-112">メンバー</span><span class="sxs-lookup"><span data-stu-id="9cda4-112">Member</span></span> | <span data-ttu-id="9cda4-113">種類</span><span class="sxs-lookup"><span data-stu-id="9cda4-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="9cda4-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="9cda4-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="9cda4-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="9cda4-115">Member</span></span> |
| [<span data-ttu-id="9cda4-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="9cda4-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="9cda4-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="9cda4-117">Member</span></span> |
| [<span data-ttu-id="9cda4-118">EventType</span><span class="sxs-lookup"><span data-stu-id="9cda4-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="9cda4-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="9cda4-119">Member</span></span> |
| [<span data-ttu-id="9cda4-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="9cda4-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="9cda4-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="9cda4-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="9cda4-122">名前空間</span><span class="sxs-lookup"><span data-stu-id="9cda4-122">Namespaces</span></span>

<span data-ttu-id="9cda4-123">[context](office.context.md):Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="9cda4-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="9cda4-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype):ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="9cda4-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="9cda4-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="9cda4-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="9cda4-126">AsyncResultStatus: 文字列</span><span class="sxs-lookup"><span data-stu-id="9cda4-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="9cda4-127">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="9cda4-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="9cda4-128">種類:</span><span class="sxs-lookup"><span data-stu-id="9cda4-128">Type:</span></span>

*   <span data-ttu-id="9cda4-129">文字列</span><span class="sxs-lookup"><span data-stu-id="9cda4-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9cda4-130">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="9cda4-130">Properties:</span></span>

|<span data-ttu-id="9cda4-131">名前</span><span class="sxs-lookup"><span data-stu-id="9cda4-131">Name</span></span>| <span data-ttu-id="9cda4-132">種類</span><span class="sxs-lookup"><span data-stu-id="9cda4-132">Type</span></span>| <span data-ttu-id="9cda4-133">説明</span><span class="sxs-lookup"><span data-stu-id="9cda4-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="9cda4-134">文字列</span><span class="sxs-lookup"><span data-stu-id="9cda4-134">String</span></span>|<span data-ttu-id="9cda4-135">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="9cda4-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="9cda4-136">文字列</span><span class="sxs-lookup"><span data-stu-id="9cda4-136">String</span></span>|<span data-ttu-id="9cda4-137">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="9cda4-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9cda4-138">要件</span><span class="sxs-lookup"><span data-stu-id="9cda4-138">Requirements</span></span>

|<span data-ttu-id="9cda4-139">要件</span><span class="sxs-lookup"><span data-stu-id="9cda4-139">Requirement</span></span>| <span data-ttu-id="9cda4-140">値</span><span class="sxs-lookup"><span data-stu-id="9cda4-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="9cda4-141">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="9cda4-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9cda4-142">1.0</span><span class="sxs-lookup"><span data-stu-id="9cda4-142">1.0</span></span>|
|[<span data-ttu-id="9cda4-143">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9cda4-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9cda4-144">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="9cda4-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="9cda4-145">CoercionType: 文字列</span><span class="sxs-lookup"><span data-stu-id="9cda4-145">CoercionType :String</span></span>

<span data-ttu-id="9cda4-146">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="9cda4-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9cda4-147">種類:</span><span class="sxs-lookup"><span data-stu-id="9cda4-147">Type:</span></span>

*   <span data-ttu-id="9cda4-148">文字列</span><span class="sxs-lookup"><span data-stu-id="9cda4-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9cda4-149">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="9cda4-149">Properties:</span></span>

|<span data-ttu-id="9cda4-150">名前</span><span class="sxs-lookup"><span data-stu-id="9cda4-150">Name</span></span>| <span data-ttu-id="9cda4-151">種類</span><span class="sxs-lookup"><span data-stu-id="9cda4-151">Type</span></span>| <span data-ttu-id="9cda4-152">説明</span><span class="sxs-lookup"><span data-stu-id="9cda4-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="9cda4-153">文字列</span><span class="sxs-lookup"><span data-stu-id="9cda4-153">String</span></span>|<span data-ttu-id="9cda4-154">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="9cda4-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="9cda4-155">文字列</span><span class="sxs-lookup"><span data-stu-id="9cda4-155">String</span></span>|<span data-ttu-id="9cda4-156">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="9cda4-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9cda4-157">要件</span><span class="sxs-lookup"><span data-stu-id="9cda4-157">Requirements</span></span>

|<span data-ttu-id="9cda4-158">要件</span><span class="sxs-lookup"><span data-stu-id="9cda4-158">Requirement</span></span>| <span data-ttu-id="9cda4-159">値</span><span class="sxs-lookup"><span data-stu-id="9cda4-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="9cda4-160">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="9cda4-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9cda4-161">1.0</span><span class="sxs-lookup"><span data-stu-id="9cda4-161">1.0</span></span>|
|[<span data-ttu-id="9cda4-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9cda4-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9cda4-163">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="9cda4-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="9cda4-164">イベントの種類: 文字列</span><span class="sxs-lookup"><span data-stu-id="9cda4-164">EventType :String</span></span>

<span data-ttu-id="9cda4-165">イベント ハンドラに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="9cda4-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="9cda4-166">種類:</span><span class="sxs-lookup"><span data-stu-id="9cda4-166">Type:</span></span>

*   <span data-ttu-id="9cda4-167">文字列</span><span class="sxs-lookup"><span data-stu-id="9cda4-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9cda4-168">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="9cda4-168">Properties:</span></span>

| <span data-ttu-id="9cda4-169">名前</span><span class="sxs-lookup"><span data-stu-id="9cda4-169">Name</span></span> | <span data-ttu-id="9cda4-170">種類</span><span class="sxs-lookup"><span data-stu-id="9cda4-170">Type</span></span> | <span data-ttu-id="9cda4-171">説明</span><span class="sxs-lookup"><span data-stu-id="9cda4-171">Description</span></span> | <span data-ttu-id="9cda4-172">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="9cda4-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="9cda4-173">文字列</span><span class="sxs-lookup"><span data-stu-id="9cda4-173">String</span></span> | <span data-ttu-id="9cda4-174">選択した予定または系列の日付または時間が変更されました。</span><span class="sxs-lookup"><span data-stu-id="9cda4-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="9cda4-175">1.7</span><span class="sxs-lookup"><span data-stu-id="9cda4-175">-17</span></span> |
|`ItemChanged`| <span data-ttu-id="9cda4-176">文字列</span><span class="sxs-lookup"><span data-stu-id="9cda4-176">String</span></span> | <span data-ttu-id="9cda4-177">選択したアイテムが変更されました。</span><span class="sxs-lookup"><span data-stu-id="9cda4-177">The selected item has changed.</span></span> | <span data-ttu-id="9cda4-178">1.5</span><span class="sxs-lookup"><span data-stu-id="9cda4-178">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="9cda4-179">文字列</span><span class="sxs-lookup"><span data-stu-id="9cda4-179">String</span></span> | <span data-ttu-id="9cda4-180">選択したアイテムが変更されました。</span><span class="sxs-lookup"><span data-stu-id="9cda4-180">The selected item has changed.</span></span> | <span data-ttu-id="9cda4-181">プレビュー</span><span class="sxs-lookup"><span data-stu-id="9cda4-181">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="9cda4-182">文字列</span><span class="sxs-lookup"><span data-stu-id="9cda4-182">String</span></span> | <span data-ttu-id="9cda4-183">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="9cda4-183">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="9cda4-184">1.7</span><span class="sxs-lookup"><span data-stu-id="9cda4-184">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="9cda4-185">文字列</span><span class="sxs-lookup"><span data-stu-id="9cda4-185">String</span></span> | <span data-ttu-id="9cda4-186">選択した系列の定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="9cda4-186">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="9cda4-187">1.7</span><span class="sxs-lookup"><span data-stu-id="9cda4-187">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9cda4-188">要件</span><span class="sxs-lookup"><span data-stu-id="9cda4-188">Requirements</span></span>

|<span data-ttu-id="9cda4-189">要件</span><span class="sxs-lookup"><span data-stu-id="9cda4-189">Requirement</span></span>| <span data-ttu-id="9cda4-190">値</span><span class="sxs-lookup"><span data-stu-id="9cda4-190">Value</span></span>|
|---|---|
|[<span data-ttu-id="9cda4-191">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="9cda4-191">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9cda4-192">1.5</span><span class="sxs-lookup"><span data-stu-id="9cda4-192">1.5</span></span> |
|[<span data-ttu-id="9cda4-193">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9cda4-193">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9cda4-194">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="9cda4-194">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="9cda4-195">SourceProperty: 文字列</span><span class="sxs-lookup"><span data-stu-id="9cda4-195">SourceProperty :String</span></span>

<span data-ttu-id="9cda4-196">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="9cda4-196">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9cda4-197">種類:</span><span class="sxs-lookup"><span data-stu-id="9cda4-197">Type:</span></span>

*   <span data-ttu-id="9cda4-198">文字列</span><span class="sxs-lookup"><span data-stu-id="9cda4-198">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9cda4-199">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="9cda4-199">Properties:</span></span>

|<span data-ttu-id="9cda4-200">名前</span><span class="sxs-lookup"><span data-stu-id="9cda4-200">Name</span></span>| <span data-ttu-id="9cda4-201">種類</span><span class="sxs-lookup"><span data-stu-id="9cda4-201">Type</span></span>| <span data-ttu-id="9cda4-202">説明</span><span class="sxs-lookup"><span data-stu-id="9cda4-202">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="9cda4-203">文字列</span><span class="sxs-lookup"><span data-stu-id="9cda4-203">String</span></span>|<span data-ttu-id="9cda4-204">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="9cda4-204">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="9cda4-205">文字列</span><span class="sxs-lookup"><span data-stu-id="9cda4-205">String</span></span>|<span data-ttu-id="9cda4-206">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="9cda4-206">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9cda4-207">要件</span><span class="sxs-lookup"><span data-stu-id="9cda4-207">Requirements</span></span>

|<span data-ttu-id="9cda4-208">要件</span><span class="sxs-lookup"><span data-stu-id="9cda4-208">Requirement</span></span>| <span data-ttu-id="9cda4-209">値</span><span class="sxs-lookup"><span data-stu-id="9cda4-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="9cda4-210">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="9cda4-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9cda4-211">1.0</span><span class="sxs-lookup"><span data-stu-id="9cda4-211">1.0</span></span>|
|[<span data-ttu-id="9cda4-212">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="9cda4-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9cda4-213">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="9cda4-213">Compose or read</span></span>|