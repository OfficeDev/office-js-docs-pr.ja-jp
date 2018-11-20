 

# <a name="office"></a><span data-ttu-id="f426b-101">Office</span><span class="sxs-lookup"><span data-stu-id="f426b-101">Office</span></span>

<span data-ttu-id="f426b-p101">Office 名前空間は、すべての Office アプリケーションのアドインで使用される共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office 名前空間の完全な一覧については、「[共有 API](/javascript/api/office)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f426b-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f426b-104">要件</span><span class="sxs-lookup"><span data-stu-id="f426b-104">Requirements</span></span>

|<span data-ttu-id="f426b-105">要件</span><span class="sxs-lookup"><span data-stu-id="f426b-105">Requirement</span></span>| <span data-ttu-id="f426b-106">値</span><span class="sxs-lookup"><span data-stu-id="f426b-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="f426b-107">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f426b-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f426b-108">1.0</span><span class="sxs-lookup"><span data-stu-id="f426b-108">1.0</span></span>|
|[<span data-ttu-id="f426b-109">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f426b-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f426b-110">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f426b-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f426b-111">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="f426b-111">Members and methods</span></span>

| <span data-ttu-id="f426b-112">メンバー</span><span class="sxs-lookup"><span data-stu-id="f426b-112">Member</span></span> | <span data-ttu-id="f426b-113">種類</span><span class="sxs-lookup"><span data-stu-id="f426b-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f426b-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="f426b-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="f426b-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="f426b-115">Member</span></span> |
| [<span data-ttu-id="f426b-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="f426b-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="f426b-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="f426b-117">Member</span></span> |
| [<span data-ttu-id="f426b-118">EventType</span><span class="sxs-lookup"><span data-stu-id="f426b-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="f426b-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="f426b-119">Member</span></span> |
| [<span data-ttu-id="f426b-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="f426b-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="f426b-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="f426b-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="f426b-122">名前空間</span><span class="sxs-lookup"><span data-stu-id="f426b-122">Namespaces</span></span>

<span data-ttu-id="f426b-123">[context](office.context.md): Outlook アドイン API で使用するために、Office アドイン API のコンテキストの名前空間から共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f426b-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="f426b-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): ItemType、EntityType、AttachmentType、RecipientType、ResponseType、および ItemNotificationMessageType 列挙型が含まれます。</span><span class="sxs-lookup"><span data-stu-id="f426b-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="f426b-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="f426b-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="f426b-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="f426b-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="f426b-127">非同期呼び出しの結果を指定します。</span><span class="sxs-lookup"><span data-stu-id="f426b-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f426b-128">型:</span><span class="sxs-lookup"><span data-stu-id="f426b-128">Type:</span></span>

*   <span data-ttu-id="f426b-129">String</span><span class="sxs-lookup"><span data-stu-id="f426b-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f426b-130">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="f426b-130">Properties:</span></span>

|<span data-ttu-id="f426b-131">名前</span><span class="sxs-lookup"><span data-stu-id="f426b-131">Name</span></span>| <span data-ttu-id="f426b-132">型</span><span class="sxs-lookup"><span data-stu-id="f426b-132">Type</span></span>| <span data-ttu-id="f426b-133">説明</span><span class="sxs-lookup"><span data-stu-id="f426b-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f426b-134">String</span><span class="sxs-lookup"><span data-stu-id="f426b-134">String</span></span>|<span data-ttu-id="f426b-135">呼び出しが成功しました。</span><span class="sxs-lookup"><span data-stu-id="f426b-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f426b-136">String</span><span class="sxs-lookup"><span data-stu-id="f426b-136">String</span></span>|<span data-ttu-id="f426b-137">呼び出しが失敗しました。</span><span class="sxs-lookup"><span data-stu-id="f426b-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f426b-138">要件</span><span class="sxs-lookup"><span data-stu-id="f426b-138">Requirements</span></span>

|<span data-ttu-id="f426b-139">要件</span><span class="sxs-lookup"><span data-stu-id="f426b-139">Requirement</span></span>| <span data-ttu-id="f426b-140">値</span><span class="sxs-lookup"><span data-stu-id="f426b-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="f426b-141">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f426b-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f426b-142">1.0</span><span class="sxs-lookup"><span data-stu-id="f426b-142">1.0</span></span>|
|[<span data-ttu-id="f426b-143">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f426b-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f426b-144">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f426b-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="f426b-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="f426b-145">CoercionType :String</span></span>

<span data-ttu-id="f426b-146">呼び出されたメソッドによって返される、または設定されるデータを強制的に変換する方法を指定します。</span><span class="sxs-lookup"><span data-stu-id="f426b-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f426b-147">型:</span><span class="sxs-lookup"><span data-stu-id="f426b-147">Type:</span></span>

*   <span data-ttu-id="f426b-148">String</span><span class="sxs-lookup"><span data-stu-id="f426b-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f426b-149">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="f426b-149">Properties:</span></span>

|<span data-ttu-id="f426b-150">名前</span><span class="sxs-lookup"><span data-stu-id="f426b-150">Name</span></span>| <span data-ttu-id="f426b-151">型</span><span class="sxs-lookup"><span data-stu-id="f426b-151">Type</span></span>| <span data-ttu-id="f426b-152">説明</span><span class="sxs-lookup"><span data-stu-id="f426b-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f426b-153">String</span><span class="sxs-lookup"><span data-stu-id="f426b-153">String</span></span>|<span data-ttu-id="f426b-154">HTML 形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="f426b-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f426b-155">String</span><span class="sxs-lookup"><span data-stu-id="f426b-155">String</span></span>|<span data-ttu-id="f426b-156">テキスト形式で返されるデータを要求します。</span><span class="sxs-lookup"><span data-stu-id="f426b-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f426b-157">要件</span><span class="sxs-lookup"><span data-stu-id="f426b-157">Requirements</span></span>

|<span data-ttu-id="f426b-158">要件</span><span class="sxs-lookup"><span data-stu-id="f426b-158">Requirement</span></span>| <span data-ttu-id="f426b-159">値</span><span class="sxs-lookup"><span data-stu-id="f426b-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="f426b-160">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f426b-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f426b-161">1.0</span><span class="sxs-lookup"><span data-stu-id="f426b-161">1.0</span></span>|
|[<span data-ttu-id="f426b-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f426b-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f426b-163">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f426b-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="f426b-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="f426b-164">EventType :String</span></span>

<span data-ttu-id="f426b-165">イベント ハンドラーに関連付けられているイベントを指定します。</span><span class="sxs-lookup"><span data-stu-id="f426b-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="f426b-166">型:</span><span class="sxs-lookup"><span data-stu-id="f426b-166">Type:</span></span>

*   <span data-ttu-id="f426b-167">String</span><span class="sxs-lookup"><span data-stu-id="f426b-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f426b-168">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="f426b-168">Properties:</span></span>

| <span data-ttu-id="f426b-169">名前</span><span class="sxs-lookup"><span data-stu-id="f426b-169">Name</span></span> | <span data-ttu-id="f426b-170">型</span><span class="sxs-lookup"><span data-stu-id="f426b-170">Type</span></span> | <span data-ttu-id="f426b-171">説明</span><span class="sxs-lookup"><span data-stu-id="f426b-171">Description</span></span> | <span data-ttu-id="f426b-172">最小要件セット</span><span class="sxs-lookup"><span data-stu-id="f426b-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="f426b-173">文字列</span><span class="sxs-lookup"><span data-stu-id="f426b-173">String</span></span> | <span data-ttu-id="f426b-174">選択した予定または一連の予定の日付または時刻が変更された。</span><span class="sxs-lookup"><span data-stu-id="f426b-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="f426b-175">1.7</span><span class="sxs-lookup"><span data-stu-id="f426b-175">-17</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="f426b-176">文字列</span><span class="sxs-lookup"><span data-stu-id="f426b-176">String</span></span> | <span data-ttu-id="f426b-177">アイテムに添付ファイルが追加されたか、アイテムから添付ファイルが削除された。</span><span class="sxs-lookup"><span data-stu-id="f426b-177">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="f426b-178">プレビュー</span><span class="sxs-lookup"><span data-stu-id="f426b-178">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="f426b-179">文字列</span><span class="sxs-lookup"><span data-stu-id="f426b-179">String</span></span> | <span data-ttu-id="f426b-180">作業ウィンドウがピン留めされている間、別の Outlook アイテムが選択されて表示されている。</span><span class="sxs-lookup"><span data-stu-id="f426b-180">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="f426b-181">1.5</span><span class="sxs-lookup"><span data-stu-id="f426b-181">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="f426b-182">文字列</span><span class="sxs-lookup"><span data-stu-id="f426b-182">String</span></span> | <span data-ttu-id="f426b-183">メールボックスの Office テーマが変更された。</span><span class="sxs-lookup"><span data-stu-id="f426b-183">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="f426b-184">プレビュー</span><span class="sxs-lookup"><span data-stu-id="f426b-184">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="f426b-185">文字列</span><span class="sxs-lookup"><span data-stu-id="f426b-185">String</span></span> | <span data-ttu-id="f426b-186">選択したアイテムまたは予定の場所の受信者リストが変更された。</span><span class="sxs-lookup"><span data-stu-id="f426b-186">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="f426b-187">1.7</span><span class="sxs-lookup"><span data-stu-id="f426b-187">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="f426b-188">文字列</span><span class="sxs-lookup"><span data-stu-id="f426b-188">String</span></span> | <span data-ttu-id="f426b-189">選択した一連の予定の定期的なパターンが変更された。</span><span class="sxs-lookup"><span data-stu-id="f426b-189">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="f426b-190">1.7</span><span class="sxs-lookup"><span data-stu-id="f426b-190">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f426b-191">要件</span><span class="sxs-lookup"><span data-stu-id="f426b-191">Requirements</span></span>

|<span data-ttu-id="f426b-192">要件</span><span class="sxs-lookup"><span data-stu-id="f426b-192">Requirement</span></span>| <span data-ttu-id="f426b-193">値</span><span class="sxs-lookup"><span data-stu-id="f426b-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="f426b-194">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f426b-194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f426b-195">1.5</span><span class="sxs-lookup"><span data-stu-id="f426b-195">1.5</span></span> |
|[<span data-ttu-id="f426b-196">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f426b-196">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f426b-197">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f426b-197">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="f426b-198">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="f426b-198">SourceProperty :String</span></span>

<span data-ttu-id="f426b-199">呼び出されたメソッドによって返されるデータのソースを指定します。</span><span class="sxs-lookup"><span data-stu-id="f426b-199">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f426b-200">型:</span><span class="sxs-lookup"><span data-stu-id="f426b-200">Type:</span></span>

*   <span data-ttu-id="f426b-201">String</span><span class="sxs-lookup"><span data-stu-id="f426b-201">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f426b-202">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="f426b-202">Properties:</span></span>

|<span data-ttu-id="f426b-203">名前</span><span class="sxs-lookup"><span data-stu-id="f426b-203">Name</span></span>| <span data-ttu-id="f426b-204">型</span><span class="sxs-lookup"><span data-stu-id="f426b-204">Type</span></span>| <span data-ttu-id="f426b-205">説明</span><span class="sxs-lookup"><span data-stu-id="f426b-205">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f426b-206">String</span><span class="sxs-lookup"><span data-stu-id="f426b-206">String</span></span>|<span data-ttu-id="f426b-207">データのソースは、メッセージの本文です。</span><span class="sxs-lookup"><span data-stu-id="f426b-207">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f426b-208">String</span><span class="sxs-lookup"><span data-stu-id="f426b-208">String</span></span>|<span data-ttu-id="f426b-209">データのソースは、メッセージの件名です。</span><span class="sxs-lookup"><span data-stu-id="f426b-209">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f426b-210">要件</span><span class="sxs-lookup"><span data-stu-id="f426b-210">Requirements</span></span>

|<span data-ttu-id="f426b-211">要件</span><span class="sxs-lookup"><span data-stu-id="f426b-211">Requirement</span></span>| <span data-ttu-id="f426b-212">値</span><span class="sxs-lookup"><span data-stu-id="f426b-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="f426b-213">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f426b-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f426b-214">1.0</span><span class="sxs-lookup"><span data-stu-id="f426b-214">1.0</span></span>|
|[<span data-ttu-id="f426b-215">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f426b-215">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f426b-216">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f426b-216">Compose or read</span></span>|