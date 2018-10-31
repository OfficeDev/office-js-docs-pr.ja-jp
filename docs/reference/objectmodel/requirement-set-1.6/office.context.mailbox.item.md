
# <a name="item"></a><span data-ttu-id="3a368-101">項目</span><span class="sxs-lookup"><span data-stu-id="3a368-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="3a368-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="3a368-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="3a368-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="3a368-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-105">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-105">Requirements</span></span>

|<span data-ttu-id="3a368-106">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-106">Requirement</span></span>| <span data-ttu-id="3a368-107">値</span><span class="sxs-lookup"><span data-stu-id="3a368-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-109">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-109">1.0</span></span>|
|[<span data-ttu-id="3a368-110">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="3a368-111">Restricted</span></span>|
|[<span data-ttu-id="3a368-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-113">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="3a368-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="3a368-114">Members and methods</span></span>

| <span data-ttu-id="3a368-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-115">Member</span></span> | <span data-ttu-id="3a368-116">型</span><span class="sxs-lookup"><span data-stu-id="3a368-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="3a368-117">attachments</span><span class="sxs-lookup"><span data-stu-id="3a368-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails) | <span data-ttu-id="3a368-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-118">Member</span></span> |
| [<span data-ttu-id="3a368-119">bcc</span><span class="sxs-lookup"><span data-stu-id="3a368-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="3a368-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-120">Member</span></span> |
| [<span data-ttu-id="3a368-121">body</span><span class="sxs-lookup"><span data-stu-id="3a368-121">body</span></span>](#body-bodyjavascriptapioutlook16officebody) | <span data-ttu-id="3a368-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-122">Member</span></span> |
| [<span data-ttu-id="3a368-123">cc</span><span class="sxs-lookup"><span data-stu-id="3a368-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="3a368-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-124">Member</span></span> |
| [<span data-ttu-id="3a368-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="3a368-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="3a368-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-126">Member</span></span> |
| [<span data-ttu-id="3a368-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="3a368-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="3a368-128">Member</span><span class="sxs-lookup"><span data-stu-id="3a368-128">Member</span></span> |
| [<span data-ttu-id="3a368-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="3a368-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="3a368-130">Member</span><span class="sxs-lookup"><span data-stu-id="3a368-130">Member</span></span> |
| [<span data-ttu-id="3a368-131">end</span><span class="sxs-lookup"><span data-stu-id="3a368-131">end</span></span>](#end-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="3a368-132">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-132">Member</span></span> |
| [<span data-ttu-id="3a368-133">from</span><span class="sxs-lookup"><span data-stu-id="3a368-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="3a368-134">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-134">Member</span></span> |
| [<span data-ttu-id="3a368-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="3a368-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="3a368-136">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-136">Member</span></span> |
| [<span data-ttu-id="3a368-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="3a368-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="3a368-138">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-138">Member</span></span> |
| [<span data-ttu-id="3a368-139">itemId</span><span class="sxs-lookup"><span data-stu-id="3a368-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="3a368-140">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-140">Member</span></span> |
| [<span data-ttu-id="3a368-141">itemType</span><span class="sxs-lookup"><span data-stu-id="3a368-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) | <span data-ttu-id="3a368-142">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-142">Member</span></span> |
| [<span data-ttu-id="3a368-143">location</span><span class="sxs-lookup"><span data-stu-id="3a368-143">location</span></span>](#location-stringlocationjavascriptapioutlook16officelocation) | <span data-ttu-id="3a368-144">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-144">Member</span></span> |
| [<span data-ttu-id="3a368-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="3a368-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="3a368-146">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-146">Member</span></span> |
| [<span data-ttu-id="3a368-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="3a368-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages) | <span data-ttu-id="3a368-148">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-148">Member</span></span> |
| [<span data-ttu-id="3a368-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="3a368-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="3a368-150">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-150">Member</span></span> |
| [<span data-ttu-id="3a368-151">主催者</span><span class="sxs-lookup"><span data-stu-id="3a368-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="3a368-152">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-152">Member</span></span> |
| [<span data-ttu-id="3a368-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="3a368-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="3a368-154">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-154">Member</span></span> |
| [<span data-ttu-id="3a368-155">送り主</span><span class="sxs-lookup"><span data-stu-id="3a368-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="3a368-156">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-156">Member</span></span> |
| [<span data-ttu-id="3a368-157">開始</span><span class="sxs-lookup"><span data-stu-id="3a368-157">start</span></span>](#start-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="3a368-158">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-158">Member</span></span> |
| [<span data-ttu-id="3a368-159">件名</span><span class="sxs-lookup"><span data-stu-id="3a368-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook16officesubject) | <span data-ttu-id="3a368-160">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-160">Member</span></span> |
| [<span data-ttu-id="3a368-161">宛先</span><span class="sxs-lookup"><span data-stu-id="3a368-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="3a368-162">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-162">Member</span></span> |
| [<span data-ttu-id="3a368-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="3a368-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="3a368-164">メソッド</span><span class="sxs-lookup"><span data-stu-id="3a368-164">Method</span></span> |
| [<span data-ttu-id="3a368-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="3a368-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="3a368-166">メソッド</span><span class="sxs-lookup"><span data-stu-id="3a368-166">Method</span></span> |
| [<span data-ttu-id="3a368-167">終了</span><span class="sxs-lookup"><span data-stu-id="3a368-167">close</span></span>](#close) | <span data-ttu-id="3a368-168">メソッド</span><span class="sxs-lookup"><span data-stu-id="3a368-168">Method</span></span> |
| [<span data-ttu-id="3a368-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="3a368-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="3a368-170">メソッド</span><span class="sxs-lookup"><span data-stu-id="3a368-170">Method</span></span> |
| [<span data-ttu-id="3a368-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="3a368-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="3a368-172">メソッド</span><span class="sxs-lookup"><span data-stu-id="3a368-172">Method</span></span> |
| [<span data-ttu-id="3a368-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="3a368-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="3a368-174">メソッド</span><span class="sxs-lookup"><span data-stu-id="3a368-174">Method</span></span> |
| [<span data-ttu-id="3a368-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="3a368-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="3a368-176">メソッド</span><span class="sxs-lookup"><span data-stu-id="3a368-176">Method</span></span> |
| [<span data-ttu-id="3a368-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="3a368-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="3a368-178">メソッド</span><span class="sxs-lookup"><span data-stu-id="3a368-178">Method</span></span> |
| [<span data-ttu-id="3a368-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="3a368-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="3a368-180">メソッド</span><span class="sxs-lookup"><span data-stu-id="3a368-180">Method</span></span> |
| [<span data-ttu-id="3a368-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="3a368-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="3a368-182">メソッド</span><span class="sxs-lookup"><span data-stu-id="3a368-182">Method</span></span> |
| [<span data-ttu-id="3a368-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="3a368-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="3a368-184">メソッド</span><span class="sxs-lookup"><span data-stu-id="3a368-184">Method</span></span> |
| [<span data-ttu-id="3a368-185">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="3a368-185">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="3a368-186">メソッド</span><span class="sxs-lookup"><span data-stu-id="3a368-186">Method</span></span> |
| [<span data-ttu-id="3a368-187">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="3a368-187">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="3a368-188">メソッド</span><span class="sxs-lookup"><span data-stu-id="3a368-188">Method</span></span> |
| [<span data-ttu-id="3a368-189">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="3a368-189">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="3a368-190">メソッド</span><span class="sxs-lookup"><span data-stu-id="3a368-190">Method</span></span> |
| [<span data-ttu-id="3a368-191">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="3a368-191">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="3a368-192">メソッド</span><span class="sxs-lookup"><span data-stu-id="3a368-192">Method</span></span> |
| [<span data-ttu-id="3a368-193">saveAsync</span><span class="sxs-lookup"><span data-stu-id="3a368-193">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="3a368-194">メソッド</span><span class="sxs-lookup"><span data-stu-id="3a368-194">Method</span></span> |
| [<span data-ttu-id="3a368-195">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="3a368-195">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="3a368-196">メソッド</span><span class="sxs-lookup"><span data-stu-id="3a368-196">Method</span></span> |

### <a name="example"></a><span data-ttu-id="3a368-197">例</span><span class="sxs-lookup"><span data-stu-id="3a368-197">Example</span></span>

<span data-ttu-id="3a368-198">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="3a368-198">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a><span data-ttu-id="3a368-199">メンバー</span><span class="sxs-lookup"><span data-stu-id="3a368-199">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="3a368-200">添付ファイル：配列.<[添付ファイルの詳細 ](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="3a368-200">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="3a368-p102">項目の添付ファイルの配列を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="3a368-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="3a368-203">潜在的なセキュリティ問題により特定の種類のファイルは、Outlookでブロックされ、したがって戻ってきません。</span><span class="sxs-lookup"><span data-stu-id="3a368-203">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="3a368-204">詳細については、[「Outlook でブロックされた添付ファイル」](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)をご覧下さい。</span><span class="sxs-lookup"><span data-stu-id="3a368-204">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-205">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-205">Type:</span></span>

*   <span data-ttu-id="3a368-206">配列.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="3a368-206">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-207">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-207">Requirements</span></span>

|<span data-ttu-id="3a368-208">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-208">Requirement</span></span>| <span data-ttu-id="3a368-209">値</span><span class="sxs-lookup"><span data-stu-id="3a368-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-210">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-211">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-211">1.0</span></span>|
|[<span data-ttu-id="3a368-212">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-213">ReadItem</span></span>|
|[<span data-ttu-id="3a368-214">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-215">読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-216">例</span><span class="sxs-lookup"><span data-stu-id="3a368-216">Example</span></span>

<span data-ttu-id="3a368-217">次のコードでは、現在の項目にあるすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="3a368-217">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="3a368-218">bcc:[受信者](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3a368-218">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="3a368-219">メッセージの BCC (ブラインド カーボン コピー) 列上の 受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="3a368-219">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="3a368-220">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="3a368-220">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-221">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-221">Type:</span></span>

*   [<span data-ttu-id="3a368-222">受信者</span><span class="sxs-lookup"><span data-stu-id="3a368-222">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="3a368-223">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-223">Requirements</span></span>

|<span data-ttu-id="3a368-224">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-224">Requirement</span></span>| <span data-ttu-id="3a368-225">値</span><span class="sxs-lookup"><span data-stu-id="3a368-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-226">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-227">1.1</span><span class="sxs-lookup"><span data-stu-id="3a368-227">1.1</span></span>|
|[<span data-ttu-id="3a368-228">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-229">ReadItem</span></span>|
|[<span data-ttu-id="3a368-230">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-231">新規作成</span><span class="sxs-lookup"><span data-stu-id="3a368-231">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-232">例</span><span class="sxs-lookup"><span data-stu-id="3a368-232">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="3a368-233">本文:[本文](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="3a368-233">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="3a368-234">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="3a368-234">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-235">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-235">Type:</span></span>

*   [<span data-ttu-id="3a368-236">本文</span><span class="sxs-lookup"><span data-stu-id="3a368-236">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="3a368-237">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-237">Requirements</span></span>

|<span data-ttu-id="3a368-238">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-238">Requirement</span></span>| <span data-ttu-id="3a368-239">値</span><span class="sxs-lookup"><span data-stu-id="3a368-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-240">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-241">1.1</span><span class="sxs-lookup"><span data-stu-id="3a368-241">1.1</span></span>|
|[<span data-ttu-id="3a368-242">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-243">ReadItem</span></span>|
|[<span data-ttu-id="3a368-244">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-245">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-245">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="3a368-246">cc: 配列。 <[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3a368-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="3a368-247">メッセージの CC (カーボン コピー) 受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="3a368-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="3a368-248">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="3a368-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3a368-249">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="3a368-249">Read mode</span></span>

<span data-ttu-id="3a368-p106">`cc`プロパティは、メッセージの**CC**列にある各受信者一覧の`EmailAddressDetails`オブジェクトを含む配列を返します。コレクションは最大100個のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="3a368-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="3a368-252">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="3a368-252">Compose mode</span></span>

<span data-ttu-id="3a368-253">`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-253">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-254">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-254">Type:</span></span>

*   <span data-ttu-id="3a368-255">配列 。<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3a368-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-256">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-256">Requirements</span></span>

|<span data-ttu-id="3a368-257">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-257">Requirement</span></span>| <span data-ttu-id="3a368-258">値</span><span class="sxs-lookup"><span data-stu-id="3a368-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-259">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-260">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-260">1.0</span></span>|
|[<span data-ttu-id="3a368-261">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-261">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-262">ReadItem</span></span>|
|[<span data-ttu-id="3a368-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-263">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-264">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-264">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-265">例</span><span class="sxs-lookup"><span data-stu-id="3a368-265">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="3a368-266">（空白が可能）conversationId：文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-266">(nullable) conversationId :String</span></span>

<span data-ttu-id="3a368-267">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="3a368-267">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="3a368-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="3a368-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="3a368-p108">作成フォームの新しいアイテムに対してこのプロパティの Null を取得します。ユーザーが件名を設定し項目を保存する場合、`conversationId`プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-272">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-272">Type:</span></span>

*   <span data-ttu-id="3a368-273">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-273">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-274">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-274">Requirements</span></span>

|<span data-ttu-id="3a368-275">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-275">Requirement</span></span>| <span data-ttu-id="3a368-276">値</span><span class="sxs-lookup"><span data-stu-id="3a368-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-277">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-277">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-278">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-278">1.0</span></span>|
|[<span data-ttu-id="3a368-279">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-279">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-280">ReadItem</span></span>|
|[<span data-ttu-id="3a368-281">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-281">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-282">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-282">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="3a368-283">dateTimeCreated: 日付</span><span class="sxs-lookup"><span data-stu-id="3a368-283">dateTimeCreated :Date</span></span>

<span data-ttu-id="3a368-p109">アイテムが作成された日時を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="3a368-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-286">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-286">Type:</span></span>

*   <span data-ttu-id="3a368-287">日付</span><span class="sxs-lookup"><span data-stu-id="3a368-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-288">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-288">Requirements</span></span>

|<span data-ttu-id="3a368-289">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-289">Requirement</span></span>| <span data-ttu-id="3a368-290">値</span><span class="sxs-lookup"><span data-stu-id="3a368-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-291">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-292">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-292">1.0</span></span>|
|[<span data-ttu-id="3a368-293">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-294">ReadItem</span></span>|
|[<span data-ttu-id="3a368-295">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-296">読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-297">例</span><span class="sxs-lookup"><span data-stu-id="3a368-297">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="3a368-298">dateTimeModified: 日付</span><span class="sxs-lookup"><span data-stu-id="3a368-298">dateTimeModified :Date</span></span>

<span data-ttu-id="3a368-p110">アイテムが最後に変更された日時を取得します。読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="3a368-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="3a368-301">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="3a368-301">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-302">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-302">Type:</span></span>

*   <span data-ttu-id="3a368-303">日付</span><span class="sxs-lookup"><span data-stu-id="3a368-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-304">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-304">Requirements</span></span>

|<span data-ttu-id="3a368-305">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-305">Requirement</span></span>| <span data-ttu-id="3a368-306">値</span><span class="sxs-lookup"><span data-stu-id="3a368-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-307">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-308">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-308">1.0</span></span>|
|[<span data-ttu-id="3a368-309">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-309">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-310">ReadItem</span></span>|
|[<span data-ttu-id="3a368-311">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-311">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-312">読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-313">例</span><span class="sxs-lookup"><span data-stu-id="3a368-313">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="3a368-314">end:日付|[時間](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="3a368-314">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="3a368-315">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="3a368-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="3a368-p111">`end`プロパティは、協定世界時 (UTC) 形式の時刻値として表示されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="3a368-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3a368-318">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="3a368-318">Read mode</span></span>

<span data-ttu-id="3a368-319">`end`プロパティは`Date`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-319">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="3a368-320">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="3a368-320">Compose mode</span></span>

<span data-ttu-id="3a368-321">`end`プロパティは`Time`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="3a368-322">[ `Time.setAsync` ](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3a368-322">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-323">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-323">Type:</span></span>

*   <span data-ttu-id="3a368-324">日付| [時間](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="3a368-324">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-325">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-325">Requirements</span></span>

|<span data-ttu-id="3a368-326">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-326">Requirement</span></span>| <span data-ttu-id="3a368-327">値</span><span class="sxs-lookup"><span data-stu-id="3a368-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-328">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-329">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-329">1.0</span></span>|
|[<span data-ttu-id="3a368-330">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-330">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-331">ReadItem</span></span>|
|[<span data-ttu-id="3a368-332">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-332">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-333">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-333">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-334">例</span><span class="sxs-lookup"><span data-stu-id="3a368-334">Example</span></span>

<span data-ttu-id="3a368-335">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="3a368-335">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="3a368-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="3a368-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="3a368-p112">メッセージの送信者の電子メール アドレスを取得します。読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="3a368-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="3a368-p113">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) プロパティは同一人物を表します。その場合、`from` プロパティは委任者を、送信者プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="3a368-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="3a368-341">`from` プロパティ内の`EmailAddressDetails` オブジェクトの`recipientType`  プロパティは、 `undefined`です。</span><span class="sxs-lookup"><span data-stu-id="3a368-341">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-342">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-342">Type:</span></span>

*   [<span data-ttu-id="3a368-343">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="3a368-343">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="3a368-344">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-344">Requirements</span></span>

|<span data-ttu-id="3a368-345">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-345">Requirement</span></span>| <span data-ttu-id="3a368-346">値</span><span class="sxs-lookup"><span data-stu-id="3a368-346">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-347">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-347">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-348">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-348">1.0</span></span>|
|[<span data-ttu-id="3a368-349">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-349">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-350">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-350">ReadItem</span></span>|
|[<span data-ttu-id="3a368-351">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-351">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-352">読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-352">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="3a368-353">internetMessageId: 文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-353">internetMessageId :String</span></span>

<span data-ttu-id="3a368-p114">電子メール メッセージのインターネット メッセージ 識別子を取得します。読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="3a368-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-356">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-356">Type:</span></span>

*   <span data-ttu-id="3a368-357">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-358">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-358">Requirements</span></span>

|<span data-ttu-id="3a368-359">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-359">Requirement</span></span>| <span data-ttu-id="3a368-360">値</span><span class="sxs-lookup"><span data-stu-id="3a368-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-361">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-362">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-362">1.0</span></span>|
|[<span data-ttu-id="3a368-363">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-364">ReadItem</span></span>|
|[<span data-ttu-id="3a368-365">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-366">読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-367">例</span><span class="sxs-lookup"><span data-stu-id="3a368-367">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="3a368-368">itemClass: 文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-368">itemClass :String</span></span>

<span data-ttu-id="3a368-p115">選択された項目の Exchange Web サービスの項目 クラスを取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="3a368-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="3a368-p116">`itemClass` プロパティは、選択したアイテムのメッセージ クラスを指定します。次は、メッセージまたは予定アイテムの既定のメッセージ クラスを示しています。</span><span class="sxs-lookup"><span data-stu-id="3a368-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="3a368-373">種類</span><span class="sxs-lookup"><span data-stu-id="3a368-373">Type</span></span> | <span data-ttu-id="3a368-374">説明</span><span class="sxs-lookup"><span data-stu-id="3a368-374">Description</span></span> | <span data-ttu-id="3a368-375">項目のクラス</span><span class="sxs-lookup"><span data-stu-id="3a368-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="3a368-376">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="3a368-376">Appointment items</span></span> | <span data-ttu-id="3a368-377">これらは、アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurence` の予定表アイテムです。</span><span class="sxs-lookup"><span data-stu-id="3a368-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="3a368-378">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="3a368-378">Message items</span></span> | <span data-ttu-id="3a368-379">これには、基本のメッセージ クラス として `IPM.Schedule.Meeting`  を使用する、既定のメッセージ クラス `IPM.Note`  会議出席依頼、返信および取り消しを持つ電子メール メッセージが含まれます。</span><span class="sxs-lookup"><span data-stu-id="3a368-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="3a368-380">既定のメッセージ クラスを拡張したカスタム メッセージ クラス 、たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など)を作成できます。</span><span class="sxs-lookup"><span data-stu-id="3a368-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-381">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-381">Type:</span></span>

*   <span data-ttu-id="3a368-382">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-383">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-383">Requirements</span></span>

|<span data-ttu-id="3a368-384">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-384">Requirement</span></span>| <span data-ttu-id="3a368-385">値</span><span class="sxs-lookup"><span data-stu-id="3a368-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-386">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-387">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-387">1.0</span></span>|
|[<span data-ttu-id="3a368-388">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-388">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-389">ReadItem</span></span>|
|[<span data-ttu-id="3a368-390">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-390">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-391">読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-392">例</span><span class="sxs-lookup"><span data-stu-id="3a368-392">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="3a368-393">（空白が可能） itemId ：文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-393">(nullable) itemId :String</span></span>

<span data-ttu-id="3a368-p117">現在の項目の Exchange Web サービスのアイテム識別子を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="3a368-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="3a368-396">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="3a368-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="3a368-397">`itemId` プロパティは、Outlook の Entry ID または Outlook の REST API によって使用される ID と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="3a368-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="3a368-398">この値を使用して REST API の呼び出しを行う前に [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3a368-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="3a368-399">詳細については、 [Outlook アドインから Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3a368-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="3a368-p119">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-402">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-402">Type:</span></span>

*   <span data-ttu-id="3a368-403">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-404">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-404">Requirements</span></span>

|<span data-ttu-id="3a368-405">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-405">Requirement</span></span>| <span data-ttu-id="3a368-406">値</span><span class="sxs-lookup"><span data-stu-id="3a368-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-407">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-408">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-408">1.0</span></span>|
|[<span data-ttu-id="3a368-409">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-409">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-410">ReadItem</span></span>|
|[<span data-ttu-id="3a368-411">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-411">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-412">読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-413">例</span><span class="sxs-lookup"><span data-stu-id="3a368-413">Example</span></span>

<span data-ttu-id="3a368-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="3a368-416">itemType:[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="3a368-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="3a368-417">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="3a368-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="3a368-418">`itemType`プロパティは、`ItemType`列挙値の 1 つを返します。これは`item`オブジェクト インスタンスがメッセージまたは予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="3a368-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-419">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-419">Type:</span></span>

*   [<span data-ttu-id="3a368-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="3a368-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="3a368-421">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-421">Requirements</span></span>

|<span data-ttu-id="3a368-422">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-422">Requirement</span></span>| <span data-ttu-id="3a368-423">値</span><span class="sxs-lookup"><span data-stu-id="3a368-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-424">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-425">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-425">1.0</span></span>|
|[<span data-ttu-id="3a368-426">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-427">ReadItem</span></span>|
|[<span data-ttu-id="3a368-428">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-429">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-429">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-430">例</span><span class="sxs-lookup"><span data-stu-id="3a368-430">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="3a368-431">位置: 文字列|[](/javascript/api/outlook_1_6/office.location)位置</span><span class="sxs-lookup"><span data-stu-id="3a368-431">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="3a368-432">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="3a368-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3a368-433">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="3a368-433">Read mode</span></span>

<span data-ttu-id="3a368-434">`location` プロパティは、予定の場所を含む文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-434">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="3a368-435">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="3a368-435">Compose mode</span></span>

<span data-ttu-id="3a368-436">`location` プロパティは、予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-437">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-437">Type:</span></span>

*   <span data-ttu-id="3a368-438">文字列 | [場所](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="3a368-438">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-439">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-439">Requirements</span></span>

|<span data-ttu-id="3a368-440">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-440">Requirement</span></span>| <span data-ttu-id="3a368-441">値</span><span class="sxs-lookup"><span data-stu-id="3a368-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-442">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-443">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-443">1.0</span></span>|
|[<span data-ttu-id="3a368-444">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-445">ReadItem</span></span>|
|[<span data-ttu-id="3a368-446">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-447">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-448">例</span><span class="sxs-lookup"><span data-stu-id="3a368-448">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="3a368-449">normalizedSubject: 文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-449">normalizedSubject :String</span></span>

<span data-ttu-id="3a368-p121">すべてのプレフィックス (`RE:` や `FWD:` を含む) が削除された項目の件名を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="3a368-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="3a368-p122">normalizedSubject プロパティは、電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたアイテムの件名を取得します。プレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="3a368-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-454">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-454">Type:</span></span>

*   <span data-ttu-id="3a368-455">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-455">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-456">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-456">Requirements</span></span>

|<span data-ttu-id="3a368-457">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-457">Requirement</span></span>| <span data-ttu-id="3a368-458">値</span><span class="sxs-lookup"><span data-stu-id="3a368-458">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-459">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-459">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-460">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-460">1.0</span></span>|
|[<span data-ttu-id="3a368-461">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-461">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-462">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-462">ReadItem</span></span>|
|[<span data-ttu-id="3a368-463">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-463">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-464">読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-464">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-465">例</span><span class="sxs-lookup"><span data-stu-id="3a368-465">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="3a368-466">notificationMessages:[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="3a368-466">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="3a368-467">項目の通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="3a368-467">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-468">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-468">Type:</span></span>

*   [<span data-ttu-id="3a368-469">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="3a368-469">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="3a368-470">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-470">Requirements</span></span>

|<span data-ttu-id="3a368-471">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-471">Requirement</span></span>| <span data-ttu-id="3a368-472">値</span><span class="sxs-lookup"><span data-stu-id="3a368-472">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-473">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-473">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-474">1.3</span><span class="sxs-lookup"><span data-stu-id="3a368-474">1.3</span></span>|
|[<span data-ttu-id="3a368-475">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-475">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-476">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-476">ReadItem</span></span>|
|[<span data-ttu-id="3a368-477">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-477">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-478">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-478">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="3a368-479">optionalAttendees: 配列.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3a368-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="3a368-480">イベントの任意の出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="3a368-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="3a368-481">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="3a368-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3a368-482">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="3a368-482">Read mode</span></span>

<span data-ttu-id="3a368-483">`optionalAttendees`プロパティは、会議への各任意出席者の`EmailAddressDetails`オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="3a368-484">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="3a368-484">Compose mode</span></span>

<span data-ttu-id="3a368-485">`optionalAttendees`プロパティは会議への任意出席者を取得または設定するためのメソッドを提供する`Recipients`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-486">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-486">Type:</span></span>

*   <span data-ttu-id="3a368-487">配列 。<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> |[受信者](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3a368-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-488">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-488">Requirements</span></span>

|<span data-ttu-id="3a368-489">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-489">Requirement</span></span>| <span data-ttu-id="3a368-490">値</span><span class="sxs-lookup"><span data-stu-id="3a368-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-491">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-491">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-492">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-492">1.0</span></span>|
|[<span data-ttu-id="3a368-493">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-493">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-494">ReadItem</span></span>|
|[<span data-ttu-id="3a368-495">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-495">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-496">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-496">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-497">例</span><span class="sxs-lookup"><span data-stu-id="3a368-497">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="3a368-498">開催者:[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="3a368-498">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="3a368-p124">指定の会議の開催者の電子メール アドレスを取得します。読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="3a368-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-501">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-501">Type:</span></span>

*   [<span data-ttu-id="3a368-502">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="3a368-502">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="3a368-503">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-503">Requirements</span></span>

|<span data-ttu-id="3a368-504">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-504">Requirement</span></span>| <span data-ttu-id="3a368-505">値</span><span class="sxs-lookup"><span data-stu-id="3a368-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-506">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-507">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-507">1.0</span></span>|
|[<span data-ttu-id="3a368-508">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-509">ReadItem</span></span>|
|[<span data-ttu-id="3a368-510">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-511">読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-511">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-512">例</span><span class="sxs-lookup"><span data-stu-id="3a368-512">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="3a368-513">requiredAttendees: 配列 。<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_6/office.recipients)。</span><span class="sxs-lookup"><span data-stu-id="3a368-513">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="3a368-514">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="3a368-514">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="3a368-515">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="3a368-515">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3a368-516">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="3a368-516">Read mode</span></span>

<span data-ttu-id="3a368-517">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-517">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="3a368-518">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="3a368-518">Compose mode</span></span>

<span data-ttu-id="3a368-519">`requiredAttendees` プロパティは会議への必須出席者を取得または設定するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-519">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-520">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-520">Type:</span></span>

*   <span data-ttu-id="3a368-521">配列 。<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3a368-521">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-522">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-522">Requirements</span></span>

|<span data-ttu-id="3a368-523">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-523">Requirement</span></span>| <span data-ttu-id="3a368-524">値</span><span class="sxs-lookup"><span data-stu-id="3a368-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-525">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-525">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-526">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-526">1.0</span></span>|
|[<span data-ttu-id="3a368-527">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-527">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-528">ReadItem</span></span>|
|[<span data-ttu-id="3a368-529">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-529">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-530">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-530">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-531">例</span><span class="sxs-lookup"><span data-stu-id="3a368-531">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="3a368-532">送信者:[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="3a368-532">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="3a368-p126">電子メール送信者のメールアドレスを取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="3a368-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="3a368-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。その場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="3a368-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="3a368-537">`sender`プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType`プロパティは、`undefined`です。</span><span class="sxs-lookup"><span data-stu-id="3a368-537">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-538">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-538">Type:</span></span>

*   [<span data-ttu-id="3a368-539">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="3a368-539">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="3a368-540">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-540">Requirements</span></span>

|<span data-ttu-id="3a368-541">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-541">Requirement</span></span>| <span data-ttu-id="3a368-542">値</span><span class="sxs-lookup"><span data-stu-id="3a368-542">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-543">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-543">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-544">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-544">1.0</span></span>|
|[<span data-ttu-id="3a368-545">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-545">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-546">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-546">ReadItem</span></span>|
|[<span data-ttu-id="3a368-547">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-547">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-548">読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-548">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-549">例</span><span class="sxs-lookup"><span data-stu-id="3a368-549">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="3a368-550">開始: 日付 | [ 時間](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="3a368-550">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="3a368-551">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="3a368-551">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="3a368-p128">`start`プロパティは、協定世界時 (UTC) 形式の時刻値として表示されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime)メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="3a368-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3a368-554">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="3a368-554">Read mode</span></span>

<span data-ttu-id="3a368-555">`start`プロパティは`Date`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-555">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="3a368-556">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="3a368-556">Compose mode</span></span>

<span data-ttu-id="3a368-557">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-557">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="3a368-558">[ `Time.setAsync` ](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3a368-558">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-559">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-559">Type:</span></span>

*   <span data-ttu-id="3a368-560">日付| [時間](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="3a368-560">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-561">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-561">Requirements</span></span>

|<span data-ttu-id="3a368-562">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-562">Requirement</span></span>| <span data-ttu-id="3a368-563">値</span><span class="sxs-lookup"><span data-stu-id="3a368-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-564">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-565">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-565">1.0</span></span>|
|[<span data-ttu-id="3a368-566">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-567">ReadItem</span></span>|
|[<span data-ttu-id="3a368-568">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-569">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-569">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-570">例</span><span class="sxs-lookup"><span data-stu-id="3a368-570">Example</span></span>

<span data-ttu-id="3a368-571">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="3a368-571">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="3a368-572">件名: 文字列 | [件名](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="3a368-572">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="3a368-573">アイテムの件名フィールドに表示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="3a368-573">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="3a368-574">`subject`プロパティは、電子メールサーバーから送信された項目の全件名を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="3a368-574">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3a368-575">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="3a368-575">Read mode</span></span>

<span data-ttu-id="3a368-p129">`subject`プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string)プロパティを使用して、行間にある`RE:`や`FW:`のなどのすべてのプレフィックスを削除した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="3a368-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="3a368-578">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="3a368-578">Compose mode</span></span>

<span data-ttu-id="3a368-579">`subject`プロパティは、件名を取得または設定するためのメソッドを提供する`Subject`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-579">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="3a368-580">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-580">Type:</span></span>

*   <span data-ttu-id="3a368-581">文字列 | [件名](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="3a368-581">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-582">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-582">Requirements</span></span>

|<span data-ttu-id="3a368-583">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-583">Requirement</span></span>| <span data-ttu-id="3a368-584">値</span><span class="sxs-lookup"><span data-stu-id="3a368-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-585">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-585">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-586">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-586">1.0</span></span>|
|[<span data-ttu-id="3a368-587">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-587">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-588">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-588">ReadItem</span></span>|
|[<span data-ttu-id="3a368-589">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-589">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-590">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-590">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="3a368-591">to: 配列。 <[ EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails) >|   [  受信者](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3a368-591">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="3a368-592">メッセージの **宛先**列にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="3a368-592">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="3a368-593">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="3a368-593">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3a368-594">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="3a368-594">Read mode</span></span>

<span data-ttu-id="3a368-p131">`to` プロパティは`EmailAddressDetails` 、メッセージの \*\* To\*\*  行にある各受信者について、 オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="3a368-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="3a368-597">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="3a368-597">Compose mode</span></span>

<span data-ttu-id="3a368-598">`to` プロパティは、メッセージの **To** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-598">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="3a368-599">型:</span><span class="sxs-lookup"><span data-stu-id="3a368-599">Type:</span></span>

*   <span data-ttu-id="3a368-600">配列 。<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3a368-600">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-601">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-601">Requirements</span></span>

|<span data-ttu-id="3a368-602">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-602">Requirement</span></span>| <span data-ttu-id="3a368-603">値</span><span class="sxs-lookup"><span data-stu-id="3a368-603">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-604">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-604">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-605">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-605">1.0</span></span>|
|[<span data-ttu-id="3a368-606">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-606">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-607">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-607">ReadItem</span></span>|
|[<span data-ttu-id="3a368-608">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-608">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-609">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-609">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-610">例</span><span class="sxs-lookup"><span data-stu-id="3a368-610">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="3a368-611">メソッド</span><span class="sxs-lookup"><span data-stu-id="3a368-611">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="3a368-612">addFileAttachmentAsync(uri, attachmentName, [オプション], [コールバック])</span><span class="sxs-lookup"><span data-stu-id="3a368-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="3a368-613">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="3a368-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="3a368-614">`addFileAttachmentAsync`メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内の項目に添付します。</span><span class="sxs-lookup"><span data-stu-id="3a368-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="3a368-615">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="3a368-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a368-616">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="3a368-616">Parameters:</span></span>

|<span data-ttu-id="3a368-617">名前</span><span class="sxs-lookup"><span data-stu-id="3a368-617">Name</span></span>| <span data-ttu-id="3a368-618">種類</span><span class="sxs-lookup"><span data-stu-id="3a368-618">Type</span></span>| <span data-ttu-id="3a368-619">属性</span><span class="sxs-lookup"><span data-stu-id="3a368-619">Attributes</span></span>| <span data-ttu-id="3a368-620">説明</span><span class="sxs-lookup"><span data-stu-id="3a368-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="3a368-621">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-621">String</span></span>||<span data-ttu-id="3a368-p132">メッセージまたは予定に添付するファイルの場所を示す URIです。 2048 文字以内で入力してください。</span><span class="sxs-lookup"><span data-stu-id="3a368-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="3a368-624">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-624">String</span></span>||<span data-ttu-id="3a368-p133">アップロード中に表示される添付ファイルがそのファイルの名前です。 255 文字以内で入力してください。</span><span class="sxs-lookup"><span data-stu-id="3a368-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="3a368-627">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="3a368-627">Object</span></span>| <span data-ttu-id="3a368-628">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-628">&lt;optional&gt;</span></span>|<span data-ttu-id="3a368-629">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="3a368-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="3a368-630">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="3a368-630">Object</span></span> | <span data-ttu-id="3a368-631">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-631">&lt;optional&gt;</span></span> | <span data-ttu-id="3a368-632">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="3a368-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="3a368-633">ブール値</span><span class="sxs-lookup"><span data-stu-id="3a368-633">Boolean</span></span> | <span data-ttu-id="3a368-634">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-634">&lt;optional&gt;</span></span> | <span data-ttu-id="3a368-635">`true`の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="3a368-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="3a368-636">関数</span><span class="sxs-lookup"><span data-stu-id="3a368-636">function</span></span>| <span data-ttu-id="3a368-637">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-637">&lt;optional&gt;</span></span>|<span data-ttu-id="3a368-638">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="3a368-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="3a368-639">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="3a368-640">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="3a368-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="3a368-641">エラー</span><span class="sxs-lookup"><span data-stu-id="3a368-641">Errors</span></span>

| <span data-ttu-id="3a368-642">エラー コード</span><span class="sxs-lookup"><span data-stu-id="3a368-642">Error code</span></span> | <span data-ttu-id="3a368-643">説明</span><span class="sxs-lookup"><span data-stu-id="3a368-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="3a368-644">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="3a368-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="3a368-645">許可されていない拡張子付きの添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="3a368-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="3a368-646">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="3a368-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3a368-647">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-647">Requirements</span></span>

|<span data-ttu-id="3a368-648">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-648">Requirement</span></span>| <span data-ttu-id="3a368-649">値</span><span class="sxs-lookup"><span data-stu-id="3a368-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-650">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-651">1.1</span><span class="sxs-lookup"><span data-stu-id="3a368-651">1.1</span></span>|
|[<span data-ttu-id="3a368-652">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3a368-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="3a368-654">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-655">新規作成</span><span class="sxs-lookup"><span data-stu-id="3a368-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="3a368-656">例</span><span class="sxs-lookup"><span data-stu-id="3a368-656">Examples</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="3a368-657">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="3a368-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        
      }
    );
  }
);
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="3a368-658">addItemAttachmentAsync (itemd、attachmentName、[オプション]、[コールバック])</span><span class="sxs-lookup"><span data-stu-id="3a368-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="3a368-659">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="3a368-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="3a368-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つ項目を新規作成フォーム内の項目に添付します。コールバック メソッドを指定する場合、`asyncResult` というパラメータがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、または項目を添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメータを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="3a368-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="3a368-663">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="3a368-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="3a368-664">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドは項目を、編集中以外の項目に添付できますが、これはサポートされておらず、推奨されていません。</span><span class="sxs-lookup"><span data-stu-id="3a368-664">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a368-665">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="3a368-665">Parameters:</span></span>

|<span data-ttu-id="3a368-666">名前</span><span class="sxs-lookup"><span data-stu-id="3a368-666">Name</span></span>| <span data-ttu-id="3a368-667">種類</span><span class="sxs-lookup"><span data-stu-id="3a368-667">Type</span></span>| <span data-ttu-id="3a368-668">属性</span><span class="sxs-lookup"><span data-stu-id="3a368-668">Attributes</span></span>| <span data-ttu-id="3a368-669">説明</span><span class="sxs-lookup"><span data-stu-id="3a368-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="3a368-670">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-670">String</span></span>||<span data-ttu-id="3a368-p135">添付するアイテムの Exchange 識別子です。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="3a368-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="3a368-673">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-673">String</span></span>||<span data-ttu-id="3a368-p136">添付するアイテムの件名です。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="3a368-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="3a368-676">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="3a368-676">Object</span></span>| <span data-ttu-id="3a368-677">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-677">&lt;optional&gt;</span></span>|<span data-ttu-id="3a368-678">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="3a368-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="3a368-679">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="3a368-679">Object</span></span>| <span data-ttu-id="3a368-680">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-680">&lt;optional&gt;</span></span>|<span data-ttu-id="3a368-681">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="3a368-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="3a368-682">関数</span><span class="sxs-lookup"><span data-stu-id="3a368-682">function</span></span>| <span data-ttu-id="3a368-683">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-683">&lt;optional&gt;</span></span>|<span data-ttu-id="3a368-684">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="3a368-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="3a368-685">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="3a368-686">添付ファイルの追加に失敗した場合、 エラーの説明を提供する`asyncResult` オブジェクトが `Error` オブジェクトに含まれます。</span><span class="sxs-lookup"><span data-stu-id="3a368-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="3a368-687">エラー</span><span class="sxs-lookup"><span data-stu-id="3a368-687">Errors</span></span>

| <span data-ttu-id="3a368-688">エラー コード</span><span class="sxs-lookup"><span data-stu-id="3a368-688">Error code</span></span> | <span data-ttu-id="3a368-689">説明</span><span class="sxs-lookup"><span data-stu-id="3a368-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="3a368-690">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="3a368-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3a368-691">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-691">Requirements</span></span>

|<span data-ttu-id="3a368-692">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-692">Requirement</span></span>| <span data-ttu-id="3a368-693">値</span><span class="sxs-lookup"><span data-stu-id="3a368-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-694">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-695">1.1</span><span class="sxs-lookup"><span data-stu-id="3a368-695">1.1</span></span>|
|[<span data-ttu-id="3a368-696">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-696">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3a368-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="3a368-698">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-698">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-699">新規作成</span><span class="sxs-lookup"><span data-stu-id="3a368-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-700">例</span><span class="sxs-lookup"><span data-stu-id="3a368-700">Example</span></span>

<span data-ttu-id="3a368-701">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="3a368-702">閉じる()</span><span class="sxs-lookup"><span data-stu-id="3a368-702">close()</span></span>

<span data-ttu-id="3a368-703">新規作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="3a368-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="3a368-p137">`close`メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="3a368-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="3a368-706">Outlook on the webでは、項目が予定で、`saveAsync`を用いて事前に保存されている場合、項目が最後に保存されてから何も変更されていない場合でも、ユーザーに対して保存、破棄またはキャンセルするよう求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="3a368-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="3a368-707">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close`メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="3a368-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-708">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-708">Requirements</span></span>

|<span data-ttu-id="3a368-709">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-709">Requirement</span></span>| <span data-ttu-id="3a368-710">値</span><span class="sxs-lookup"><span data-stu-id="3a368-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-711">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-712">1.3</span><span class="sxs-lookup"><span data-stu-id="3a368-712">1.3</span></span>|
|[<span data-ttu-id="3a368-713">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-713">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-714">制限あり</span><span class="sxs-lookup"><span data-stu-id="3a368-714">Restricted</span></span>|
|[<span data-ttu-id="3a368-715">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-715">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-716">新規作成</span><span class="sxs-lookup"><span data-stu-id="3a368-716">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="3a368-717">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="3a368-717">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="3a368-718">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="3a368-719">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="3a368-719">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3a368-720">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-720">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="3a368-721">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="3a368-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="3a368-p138">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="3a368-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a368-725">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="3a368-725">Parameters:</span></span>

| <span data-ttu-id="3a368-726">名前</span><span class="sxs-lookup"><span data-stu-id="3a368-726">Name</span></span> | <span data-ttu-id="3a368-727">種類</span><span class="sxs-lookup"><span data-stu-id="3a368-727">Type</span></span> | <span data-ttu-id="3a368-728">属性</span><span class="sxs-lookup"><span data-stu-id="3a368-728">Attributes</span></span> | <span data-ttu-id="3a368-729">説明</span><span class="sxs-lookup"><span data-stu-id="3a368-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="3a368-730">文字列 |オブジェクト</span><span class="sxs-lookup"><span data-stu-id="3a368-730">String &#124; Object</span></span>| |<span data-ttu-id="3a368-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="3a368-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="3a368-733">**または**</span><span class="sxs-lookup"><span data-stu-id="3a368-733">**OR**</span></span><br/><span data-ttu-id="3a368-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクトです。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="3a368-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="3a368-736">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-736">String</span></span> | <span data-ttu-id="3a368-737">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-737">&lt;optional&gt;</span></span> | <span data-ttu-id="3a368-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="3a368-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="3a368-740">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="3a368-741">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-741">&lt;optional&gt;</span></span> | <span data-ttu-id="3a368-742">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="3a368-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="3a368-743">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-743">String</span></span> | | <span data-ttu-id="3a368-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` でなければいけません。</span><span class="sxs-lookup"><span data-stu-id="3a368-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="3a368-746">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-746">String</span></span> | | <span data-ttu-id="3a368-747">添付ファイル名を含む文字列です。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="3a368-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="3a368-748">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-748">String</span></span> | | <span data-ttu-id="3a368-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URIです。</span><span class="sxs-lookup"><span data-stu-id="3a368-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="3a368-751">ブール値</span><span class="sxs-lookup"><span data-stu-id="3a368-751">Boolean</span></span> | | <span data-ttu-id="3a368-p144">`type`が`file`に設定されている場合にのみ使用されます。`true`の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="3a368-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="3a368-754">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-754">String</span></span> | | <span data-ttu-id="3a368-p145">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムIDです。最大長が 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="3a368-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="3a368-758">関数</span><span class="sxs-lookup"><span data-stu-id="3a368-758">function</span></span> | <span data-ttu-id="3a368-759">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-759">&lt;optional&gt;</span></span> | <span data-ttu-id="3a368-760">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである シングル パラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3a368-761">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-761">Requirements</span></span>

|<span data-ttu-id="3a368-762">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-762">Requirement</span></span>| <span data-ttu-id="3a368-763">値</span><span class="sxs-lookup"><span data-stu-id="3a368-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-764">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-765">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-765">1.0</span></span>|
|[<span data-ttu-id="3a368-766">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-766">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-767">ReadItem</span></span>|
|[<span data-ttu-id="3a368-768">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-768">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-769">読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="3a368-770">例</span><span class="sxs-lookup"><span data-stu-id="3a368-770">Examples</span></span>

<span data-ttu-id="3a368-771">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="3a368-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="3a368-772">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="3a368-772">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="3a368-773">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="3a368-773">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="3a368-774">本文と添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="3a368-774">Reply with a body and a file attachment.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="3a368-775">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="3a368-775">Reply with a body and an item attachment.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="3a368-776">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="3a368-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="3a368-777">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="3a368-777">displayReplyForm(formData)</span></span>

<span data-ttu-id="3a368-778">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む返信フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="3a368-779">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="3a368-779">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3a368-780">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-780">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="3a368-781">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="3a368-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="3a368-p146">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="3a368-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a368-785">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="3a368-785">Parameters:</span></span>

| <span data-ttu-id="3a368-786">名前</span><span class="sxs-lookup"><span data-stu-id="3a368-786">Name</span></span> | <span data-ttu-id="3a368-787">種類</span><span class="sxs-lookup"><span data-stu-id="3a368-787">Type</span></span> | <span data-ttu-id="3a368-788">属性</span><span class="sxs-lookup"><span data-stu-id="3a368-788">Attributes</span></span> | <span data-ttu-id="3a368-789">説明</span><span class="sxs-lookup"><span data-stu-id="3a368-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="3a368-790">文字列 |オブジェクト</span><span class="sxs-lookup"><span data-stu-id="3a368-790">String &#124; Object</span></span>| | <span data-ttu-id="3a368-p147">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="3a368-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="3a368-793">**または**</span><span class="sxs-lookup"><span data-stu-id="3a368-793">**OR**</span></span><br/><span data-ttu-id="3a368-p148">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクトです。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="3a368-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="3a368-796">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-796">String</span></span> | <span data-ttu-id="3a368-797">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-797">&lt;optional&gt;</span></span> | <span data-ttu-id="3a368-p149">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="3a368-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="3a368-800">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="3a368-801">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-801">&lt;optional&gt;</span></span> | <span data-ttu-id="3a368-802">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="3a368-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="3a368-803">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-803">String</span></span> | | <span data-ttu-id="3a368-p150">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` でなければいけません。</span><span class="sxs-lookup"><span data-stu-id="3a368-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="3a368-806">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-806">String</span></span> | | <span data-ttu-id="3a368-807">添付ファイル名を含む文字列です。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="3a368-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="3a368-808">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-808">String</span></span> | | <span data-ttu-id="3a368-p151">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URIです。</span><span class="sxs-lookup"><span data-stu-id="3a368-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="3a368-811">ブール値</span><span class="sxs-lookup"><span data-stu-id="3a368-811">Boolean</span></span> | | <span data-ttu-id="3a368-p152">`type`が`file`に設定されている場合にのみ使用されます。`true`の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="3a368-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="3a368-814">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-814">String</span></span> | | <span data-ttu-id="3a368-p153">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムIDです。最大長が 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="3a368-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="3a368-818">関数</span><span class="sxs-lookup"><span data-stu-id="3a368-818">function</span></span> | <span data-ttu-id="3a368-819">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-819">&lt;optional&gt;</span></span> | <span data-ttu-id="3a368-820">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである シングル パラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3a368-821">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-821">Requirements</span></span>

|<span data-ttu-id="3a368-822">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-822">Requirement</span></span>| <span data-ttu-id="3a368-823">値</span><span class="sxs-lookup"><span data-stu-id="3a368-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-824">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-824">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-825">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-825">1.0</span></span>|
|[<span data-ttu-id="3a368-826">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-826">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-827">ReadItem</span></span>|
|[<span data-ttu-id="3a368-828">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-828">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-829">読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="3a368-830">例</span><span class="sxs-lookup"><span data-stu-id="3a368-830">Examples</span></span>

<span data-ttu-id="3a368-831">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="3a368-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="3a368-832">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="3a368-832">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="3a368-833">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="3a368-833">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="3a368-834">本文と添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="3a368-834">Reply with a body and a file attachment.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="3a368-835">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="3a368-835">Reply with a body and an item attachment.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="3a368-836">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="3a368-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="3a368-837">getEntities() → {[エンティティ](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="3a368-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="3a368-838">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="3a368-838">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="3a368-839">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="3a368-839">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-840">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-840">Requirements</span></span>

|<span data-ttu-id="3a368-841">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-841">Requirement</span></span>| <span data-ttu-id="3a368-842">値</span><span class="sxs-lookup"><span data-stu-id="3a368-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-843">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-844">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-844">1.0</span></span>|
|[<span data-ttu-id="3a368-845">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-845">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-846">ReadItem</span></span>|
|[<span data-ttu-id="3a368-847">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-847">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-848">読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3a368-849">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="3a368-849">Returns:</span></span>

<span data-ttu-id="3a368-850">種類: [エンティティ](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="3a368-850">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="3a368-851">例</span><span class="sxs-lookup"><span data-stu-id="3a368-851">Example</span></span>

<span data-ttu-id="3a368-852">次の例では、現在の項目の本文内にある連絡先のエンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="3a368-852">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="3a368-853">getEntitiesByType(entityType)] → [(空白可能) {<(String|[連絡先](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[電話番号](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion)) >}</span><span class="sxs-lookup"><span data-stu-id="3a368-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="3a368-854">選択した項目で見つかった指定のエンティティ型のエンティティすべてを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="3a368-854">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="3a368-855">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="3a368-855">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a368-856">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="3a368-856">Parameters:</span></span>

|<span data-ttu-id="3a368-857">名前</span><span class="sxs-lookup"><span data-stu-id="3a368-857">Name</span></span>| <span data-ttu-id="3a368-858">種類</span><span class="sxs-lookup"><span data-stu-id="3a368-858">Type</span></span>| <span data-ttu-id="3a368-859">説明</span><span class="sxs-lookup"><span data-stu-id="3a368-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="3a368-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="3a368-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="3a368-861">EntityType 列挙値の 1 つです。</span><span class="sxs-lookup"><span data-stu-id="3a368-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a368-862">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-862">Requirements</span></span>

|<span data-ttu-id="3a368-863">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-863">Requirement</span></span>| <span data-ttu-id="3a368-864">値</span><span class="sxs-lookup"><span data-stu-id="3a368-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-865">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-866">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-866">1.0</span></span>|
|[<span data-ttu-id="3a368-867">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-867">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-868">制限あり</span><span class="sxs-lookup"><span data-stu-id="3a368-868">Restricted</span></span>|
|[<span data-ttu-id="3a368-869">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-869">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-870">読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3a368-871">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="3a368-871">Returns:</span></span>

<span data-ttu-id="3a368-872">`entityType` に渡された値が有効な `EntityType` 列挙型のメンバーでない場合、メソッドは 空白を返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="3a368-873">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-873">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="3a368-874">そうでない場合、返される配列内のオブジェクトの種類は、 `entityType` パラメータ内で要求されたエンティティの種類によって異なります。</span><span class="sxs-lookup"><span data-stu-id="3a368-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="3a368-875">このメソッドを使用する最小限のアクセス許可レベルは **制限あり** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="3a368-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="3a368-876">の値 `entityType`</span><span class="sxs-lookup"><span data-stu-id="3a368-876">Value of `entityType`</span></span> | <span data-ttu-id="3a368-877">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="3a368-877">Type of objects in returned array</span></span> | <span data-ttu-id="3a368-878">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="3a368-879">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-879">String</span></span> | <span data-ttu-id="3a368-880">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="3a368-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="3a368-881">連絡先</span><span class="sxs-lookup"><span data-stu-id="3a368-881">Contact</span></span> | <span data-ttu-id="3a368-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="3a368-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="3a368-883">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-883">String</span></span> | <span data-ttu-id="3a368-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="3a368-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="3a368-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="3a368-885">MeetingSuggestion</span></span> | <span data-ttu-id="3a368-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="3a368-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="3a368-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="3a368-887">PhoneNumber</span></span> | <span data-ttu-id="3a368-888">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="3a368-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="3a368-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="3a368-889">TaskSuggestion</span></span> | <span data-ttu-id="3a368-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="3a368-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="3a368-891">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-891">String</span></span> | <span data-ttu-id="3a368-892">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="3a368-892">**Restricted**</span></span> |

<span data-ttu-id="3a368-893">型:Array.<(文字列|[連絡先](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="3a368-893">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="3a368-894">例</span><span class="sxs-lookup"><span data-stu-id="3a368-894">Example</span></span>

<span data-ttu-id="3a368-895">次の例は、現在のアイテムの本文にある郵便アドレスを表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="3a368-895">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

```
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="3a368-896">getFilteredEntitiesByName(name)] → [(Null 許容) {<(文字列| [連絡先](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[電話番号 ](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="3a368-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="3a368-897">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="3a368-898">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="3a368-898">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3a368-899">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a368-900">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="3a368-900">Parameters:</span></span>

|<span data-ttu-id="3a368-901">名前</span><span class="sxs-lookup"><span data-stu-id="3a368-901">Name</span></span>| <span data-ttu-id="3a368-902">種類</span><span class="sxs-lookup"><span data-stu-id="3a368-902">Type</span></span>| <span data-ttu-id="3a368-903">説明</span><span class="sxs-lookup"><span data-stu-id="3a368-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="3a368-904">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-904">String</span></span>|<span data-ttu-id="3a368-905">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前です。</span><span class="sxs-lookup"><span data-stu-id="3a368-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a368-906">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-906">Requirements</span></span>

|<span data-ttu-id="3a368-907">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-907">Requirement</span></span>| <span data-ttu-id="3a368-908">値</span><span class="sxs-lookup"><span data-stu-id="3a368-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-909">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-909">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-910">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-910">1.0</span></span>|
|[<span data-ttu-id="3a368-911">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-911">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-912">ReadItem</span></span>|
|[<span data-ttu-id="3a368-913">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-913">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-914">読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3a368-915">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="3a368-915">Returns:</span></span>

<span data-ttu-id="3a368-p155">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="3a368-918">型:Array.<(文字列|[連絡先](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="3a368-918">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="3a368-919">getRegExMatches() → {オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="3a368-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="3a368-920">選択した項目内で、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="3a368-921">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="3a368-921">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3a368-p156">`getRegExMatches` メソッドは、マニフェスト XML ファイル内にある各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定された項目のプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="3a368-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="3a368-925">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="3a368-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="3a368-926">`getRegExMatches` から返されるオブジェクトには、 `fruits` および `veggies` という 2 つのプロパティがありえます。</span><span class="sxs-lookup"><span data-stu-id="3a368-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="3a368-p157">アイテムの 本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。代わりに [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="3a368-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-930">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-930">Requirements</span></span>

|<span data-ttu-id="3a368-931">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-931">Requirement</span></span>| <span data-ttu-id="3a368-932">値</span><span class="sxs-lookup"><span data-stu-id="3a368-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-933">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-934">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-934">1.0</span></span>|
|[<span data-ttu-id="3a368-935">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-935">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-936">ReadItem</span></span>|
|[<span data-ttu-id="3a368-937">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-937">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-938">読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3a368-939">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="3a368-939">Returns:</span></span>

<span data-ttu-id="3a368-p158">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクトです。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="3a368-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="3a368-942">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="3a368-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="3a368-943">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="3a368-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="3a368-944">例</span><span class="sxs-lookup"><span data-stu-id="3a368-944">Example</span></span>

<span data-ttu-id="3a368-945">次の例は、マニフェストで指定された正規表現ルールの要素`fruits`および`veggies`に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="3a368-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="3a368-946">getRegExMatchesByName(name)] → [(Null 許容) {配列.< 文字列 >}</span><span class="sxs-lookup"><span data-stu-id="3a368-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="3a368-947">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="3a368-948">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="3a368-948">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3a368-949">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="3a368-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="3a368-p159">項目の 本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、項目の本文全体を返さないようにします。`.*` などの正規表現を使用して項目の本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="3a368-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a368-952">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="3a368-952">Parameters:</span></span>

|<span data-ttu-id="3a368-953">名前</span><span class="sxs-lookup"><span data-stu-id="3a368-953">Name</span></span>| <span data-ttu-id="3a368-954">種類</span><span class="sxs-lookup"><span data-stu-id="3a368-954">Type</span></span>| <span data-ttu-id="3a368-955">説明</span><span class="sxs-lookup"><span data-stu-id="3a368-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="3a368-956">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-956">String</span></span>|<span data-ttu-id="3a368-957">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前です。</span><span class="sxs-lookup"><span data-stu-id="3a368-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a368-958">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-958">Requirements</span></span>

|<span data-ttu-id="3a368-959">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-959">Requirement</span></span>| <span data-ttu-id="3a368-960">値</span><span class="sxs-lookup"><span data-stu-id="3a368-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-961">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-961">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-962">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-962">1.0</span></span>|
|[<span data-ttu-id="3a368-963">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-963">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-964">ReadItem</span></span>|
|[<span data-ttu-id="3a368-965">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-965">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-966">読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3a368-967">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="3a368-967">Returns:</span></span>

<span data-ttu-id="3a368-968">マニフェスト XML ファイルで定義された正規表現に一致する文字列を含む配列です。</span><span class="sxs-lookup"><span data-stu-id="3a368-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="3a368-969">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="3a368-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="3a368-970">配列. < 文字列 ></span><span class="sxs-lookup"><span data-stu-id="3a368-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="3a368-971">例</span><span class="sxs-lookup"><span data-stu-id="3a368-971">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="3a368-972">getSelectedDataAsync(coercionType, [オプション], コールバック) → {文字列}</span><span class="sxs-lookup"><span data-stu-id="3a368-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="3a368-973">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="3a368-p160">選択したデータがなく、カーソルが本文または件名にある場合、メソッドは選択したデータに対して Null を返します。本文または件名以外のフィールドが選択されている場合、メソッドは `InvalidSelection` エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a368-976">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="3a368-976">Parameters:</span></span>

|<span data-ttu-id="3a368-977">名前</span><span class="sxs-lookup"><span data-stu-id="3a368-977">Name</span></span>| <span data-ttu-id="3a368-978">種類</span><span class="sxs-lookup"><span data-stu-id="3a368-978">Type</span></span>| <span data-ttu-id="3a368-979">属性</span><span class="sxs-lookup"><span data-stu-id="3a368-979">Attributes</span></span>| <span data-ttu-id="3a368-980">説明</span><span class="sxs-lookup"><span data-stu-id="3a368-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="3a368-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="3a368-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="3a368-p161">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="3a368-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="3a368-985">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="3a368-985">Object</span></span>| <span data-ttu-id="3a368-986">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-986">&lt;optional&gt;</span></span>|<span data-ttu-id="3a368-987">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="3a368-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="3a368-988">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="3a368-988">Object</span></span>| <span data-ttu-id="3a368-989">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-989">&lt;optional&gt;</span></span>|<span data-ttu-id="3a368-990">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="3a368-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="3a368-991">関数</span><span class="sxs-lookup"><span data-stu-id="3a368-991">function</span></span>||<span data-ttu-id="3a368-992">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="3a368-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3a368-993">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data`を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="3a368-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="3a368-994">選択元のソース プロパティにアクセスするには、 `asyncResult.value.sourceProperty` を呼び出します。これは `body`    または `subject`  になります。</span><span class="sxs-lookup"><span data-stu-id="3a368-994">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a368-995">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-995">Requirements</span></span>

|<span data-ttu-id="3a368-996">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-996">Requirement</span></span>| <span data-ttu-id="3a368-997">値</span><span class="sxs-lookup"><span data-stu-id="3a368-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-998">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-998">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-999">1.2</span><span class="sxs-lookup"><span data-stu-id="3a368-999">1.2</span></span>|
|[<span data-ttu-id="3a368-1000">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-1000">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3a368-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="3a368-1002">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-1002">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-1003">新規作成</span><span class="sxs-lookup"><span data-stu-id="3a368-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="3a368-1004">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="3a368-1004">Returns:</span></span>

<span data-ttu-id="3a368-1005">`coercionType`で決定された書式設定の文字列として選択されたデータです。</span><span class="sxs-lookup"><span data-stu-id="3a368-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="3a368-1006">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="3a368-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="3a368-1007">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="3a368-1008">例</span><span class="sxs-lookup"><span data-stu-id="3a368-1008">Example</span></span>

```
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="3a368-1009">getSelectedEntities() → {[エンティティ](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="3a368-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="3a368-p163">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-p163">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="3a368-1012">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="3a368-1012">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-1013">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-1013">Requirements</span></span>

|<span data-ttu-id="3a368-1014">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-1014">Requirement</span></span>| <span data-ttu-id="3a368-1015">値</span><span class="sxs-lookup"><span data-stu-id="3a368-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-1016">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-1016">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="3a368-1017">-16</span></span> |
|[<span data-ttu-id="3a368-1018">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-1018">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-1019">ReadItem</span></span>|
|[<span data-ttu-id="3a368-1020">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-1020">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-1021">読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3a368-1022">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="3a368-1022">Returns:</span></span>

<span data-ttu-id="3a368-1023">種類: [エンティティ](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="3a368-1023">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="3a368-1024">例</span><span class="sxs-lookup"><span data-stu-id="3a368-1024">Example</span></span>

<span data-ttu-id="3a368-1025">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="3a368-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="3a368-1026">getSelectedRegExMatches() → {オブジェクト}</span><span class="sxs-lookup"><span data-stu-id="3a368-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="3a368-p164">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="3a368-1029">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="3a368-1029">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3a368-p165">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内にある各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定された項目のプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="3a368-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="3a368-1033">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="3a368-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="3a368-1034">`getRegExMatches` から返されるオブジェクトには、 `fruits` および `veggies` という 2 つのプロパティがありえます。</span><span class="sxs-lookup"><span data-stu-id="3a368-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="3a368-p166">アイテムの 本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。代わりに [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="3a368-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3a368-1038">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-1038">Requirements</span></span>

|<span data-ttu-id="3a368-1039">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-1039">Requirement</span></span>| <span data-ttu-id="3a368-1040">値</span><span class="sxs-lookup"><span data-stu-id="3a368-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-1041">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-1041">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="3a368-1042">-16</span></span> |
|[<span data-ttu-id="3a368-1043">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-1043">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-1044">ReadItem</span></span>|
|[<span data-ttu-id="3a368-1045">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-1045">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-1046">読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3a368-1047">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="3a368-1047">Returns:</span></span>

<span data-ttu-id="3a368-p167">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクトです。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="3a368-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="3a368-1050">例</span><span class="sxs-lookup"><span data-stu-id="3a368-1050">Example</span></span>

<span data-ttu-id="3a368-1051">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="3a368-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="3a368-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3a368-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="3a368-1053">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="3a368-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="3a368-p168">カスタム プロパティは、アプリケーションごと、アイテムごとにキーと値のペアとして保管されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在の項目および現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、項目上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="3a368-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a368-1057">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="3a368-1057">Parameters:</span></span>

|<span data-ttu-id="3a368-1058">名前</span><span class="sxs-lookup"><span data-stu-id="3a368-1058">Name</span></span>| <span data-ttu-id="3a368-1059">種類</span><span class="sxs-lookup"><span data-stu-id="3a368-1059">Type</span></span>| <span data-ttu-id="3a368-1060">属性</span><span class="sxs-lookup"><span data-stu-id="3a368-1060">Attributes</span></span>| <span data-ttu-id="3a368-1061">説明</span><span class="sxs-lookup"><span data-stu-id="3a368-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="3a368-1062">関数</span><span class="sxs-lookup"><span data-stu-id="3a368-1062">function</span></span>||<span data-ttu-id="3a368-1063">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="3a368-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3a368-1064">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="3a368-1065">項目からカスタム プロパティを取得、設定、削除して、サーバーにカスタム プロパティのセット バックに対する変更を保存するのに、このオブジェクトが使用できます。</span><span class="sxs-lookup"><span data-stu-id="3a368-1065">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="3a368-1066">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="3a368-1066">Object</span></span>| <span data-ttu-id="3a368-1067">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="3a368-1068">開発者は、コールバック 関数でアクセスしたいオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="3a368-1068">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="3a368-1069">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="3a368-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a368-1070">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-1070">Requirements</span></span>

|<span data-ttu-id="3a368-1071">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-1071">Requirement</span></span>| <span data-ttu-id="3a368-1072">値</span><span class="sxs-lookup"><span data-stu-id="3a368-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-1073">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-1073">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-1074">1.0以降</span><span class="sxs-lookup"><span data-stu-id="3a368-1074">1.0</span></span>|
|[<span data-ttu-id="3a368-1075">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-1075">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3a368-1076">ReadItem</span></span>|
|[<span data-ttu-id="3a368-1077">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-1077">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-1078">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="3a368-1078">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-1079">例</span><span class="sxs-lookup"><span data-stu-id="3a368-1079">Example</span></span>

<span data-ttu-id="3a368-p171">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="3a368-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="3a368-1083">removeAttachmentAsync(attachmentId、[オプション]、 [コールバック])</span><span class="sxs-lookup"><span data-stu-id="3a368-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="3a368-1084">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="3a368-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="3a368-p172">`removeAttachmentAsync`メソッドは、指定した識別子の添付ファイルを項目 から削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web アプリ とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別のウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="3a368-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a368-1089">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="3a368-1089">Parameters:</span></span>

|<span data-ttu-id="3a368-1090">名前</span><span class="sxs-lookup"><span data-stu-id="3a368-1090">Name</span></span>| <span data-ttu-id="3a368-1091">種類</span><span class="sxs-lookup"><span data-stu-id="3a368-1091">Type</span></span>| <span data-ttu-id="3a368-1092">属性</span><span class="sxs-lookup"><span data-stu-id="3a368-1092">Attributes</span></span>| <span data-ttu-id="3a368-1093">説明</span><span class="sxs-lookup"><span data-stu-id="3a368-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="3a368-1094">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-1094">String</span></span>||<span data-ttu-id="3a368-p173">削除する添付ファイルの識別子です。文字列の最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="3a368-p173">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="3a368-1097">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="3a368-1097">Object</span></span>| <span data-ttu-id="3a368-1098">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="3a368-1099">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="3a368-1099">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="3a368-1100">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="3a368-1100">Object</span></span>| <span data-ttu-id="3a368-1101">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-1101">&lt;optional&gt;</span></span>|<span data-ttu-id="3a368-1102">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="3a368-1102">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="3a368-1103">関数</span><span class="sxs-lookup"><span data-stu-id="3a368-1103">function</span></span>| <span data-ttu-id="3a368-1104">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-1104">&lt;optional&gt;</span></span>|<span data-ttu-id="3a368-1105">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="3a368-1105">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="3a368-1106">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="3a368-1106">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="3a368-1107">エラー</span><span class="sxs-lookup"><span data-stu-id="3a368-1107">Errors</span></span>

| <span data-ttu-id="3a368-1108">エラー コード</span><span class="sxs-lookup"><span data-stu-id="3a368-1108">Error code</span></span> | <span data-ttu-id="3a368-1109">説明</span><span class="sxs-lookup"><span data-stu-id="3a368-1109">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="3a368-1110">添付ファイルの識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="3a368-1110">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3a368-1111">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-1111">Requirements</span></span>

|<span data-ttu-id="3a368-1112">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-1112">Requirement</span></span>| <span data-ttu-id="3a368-1113">値</span><span class="sxs-lookup"><span data-stu-id="3a368-1113">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-1114">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-1114">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-1115">1.1</span><span class="sxs-lookup"><span data-stu-id="3a368-1115">1.1</span></span>|
|[<span data-ttu-id="3a368-1116">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-1116">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-1117">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3a368-1117">ReadWriteItem</span></span>|
|[<span data-ttu-id="3a368-1118">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-1118">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-1119">新規作成</span><span class="sxs-lookup"><span data-stu-id="3a368-1119">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-1120">例</span><span class="sxs-lookup"><span data-stu-id="3a368-1120">Example</span></span>

<span data-ttu-id="3a368-1121">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="3a368-1121">The following code removes an attachment with an identifier of '0'.</span></span>

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="3a368-1122">saveAsync ([オプション] 、コールバック)</span><span class="sxs-lookup"><span data-stu-id="3a368-1122">saveAsync([options], callback)</span></span>

<span data-ttu-id="3a368-1123">アイテムを非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="3a368-1123">Asynchronously saves an item.</span></span>

<span data-ttu-id="3a368-p174">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッド経由でアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-p174">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="3a368-1127">アドインが、WS または REST API を使用しようとして`itemId`を取得するために、新規作成モードでアイテム上の`saveAsync`を呼び出す場合、Outlook キャッシュ モードでは、アイテムがサーバーと実際に同期するまでに時間がかかる場合があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="3a368-1127">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="3a368-1128">項目が同期されるまで、 `itemId` を使用すると、エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-1128">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="3a368-p176">予定はドラフト状態にはならないため、作成モードで予定に`saveAsync`が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-p176">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="3a368-1132">次のクライアントは、新規作成モードで予定上の `saveAsync` に対して様々なふるまいをします。</span><span class="sxs-lookup"><span data-stu-id="3a368-1132">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="3a368-1133">Mac Outlook は、作成モードの会議で`saveAsync`をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="3a368-1133">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="3a368-1134">Mac Outlookの会議場で  `saveAsync` を呼びだすと、エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-1134">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="3a368-1135">作成モードの予定上で`saveAsync`が呼び出されると、Outlook on the web は常に、招待または更新を送信します。</span><span class="sxs-lookup"><span data-stu-id="3a368-1135">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a368-1136">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="3a368-1136">Parameters:</span></span>

|<span data-ttu-id="3a368-1137">名前</span><span class="sxs-lookup"><span data-stu-id="3a368-1137">Name</span></span>| <span data-ttu-id="3a368-1138">種類</span><span class="sxs-lookup"><span data-stu-id="3a368-1138">Type</span></span>| <span data-ttu-id="3a368-1139">属性</span><span class="sxs-lookup"><span data-stu-id="3a368-1139">Attributes</span></span>| <span data-ttu-id="3a368-1140">説明</span><span class="sxs-lookup"><span data-stu-id="3a368-1140">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="3a368-1141">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="3a368-1141">Object</span></span>| <span data-ttu-id="3a368-1142">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="3a368-1143">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="3a368-1143">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="3a368-1144">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="3a368-1144">Object</span></span>| <span data-ttu-id="3a368-1145">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-1145">&lt;optional&gt;</span></span>|<span data-ttu-id="3a368-1146">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="3a368-1146">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="3a368-1147">関数</span><span class="sxs-lookup"><span data-stu-id="3a368-1147">function</span></span>||<span data-ttu-id="3a368-1148">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="3a368-1148">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3a368-1149">成功すると、アイテム識別子が`asyncResult.value`プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-1149">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3a368-1150">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-1150">Requirements</span></span>

|<span data-ttu-id="3a368-1151">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-1151">Requirement</span></span>| <span data-ttu-id="3a368-1152">値</span><span class="sxs-lookup"><span data-stu-id="3a368-1152">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-1153">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-1153">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-1154">1.3</span><span class="sxs-lookup"><span data-stu-id="3a368-1154">1.3</span></span>|
|[<span data-ttu-id="3a368-1155">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-1155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-1156">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3a368-1156">ReadWriteItem</span></span>|
|[<span data-ttu-id="3a368-1157">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-1157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-1158">新規作成</span><span class="sxs-lookup"><span data-stu-id="3a368-1158">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="3a368-1159">例</span><span class="sxs-lookup"><span data-stu-id="3a368-1159">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="3a368-p178">次の例は、コールバック関数に渡される `result` パラメータの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="3a368-p178">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="3a368-1162">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="3a368-1162">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="3a368-1163">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="3a368-1163">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="3a368-p179">`setSelectedDataAsync`メソッドは、指定された文字列を項目のサブジェクトまたは本文のカーソル位置に挿入するか、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="3a368-p179">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3a368-1167">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="3a368-1167">Parameters:</span></span>

|<span data-ttu-id="3a368-1168">名前</span><span class="sxs-lookup"><span data-stu-id="3a368-1168">Name</span></span>| <span data-ttu-id="3a368-1169">種類</span><span class="sxs-lookup"><span data-stu-id="3a368-1169">Type</span></span>| <span data-ttu-id="3a368-1170">属性</span><span class="sxs-lookup"><span data-stu-id="3a368-1170">Attributes</span></span>| <span data-ttu-id="3a368-1171">説明</span><span class="sxs-lookup"><span data-stu-id="3a368-1171">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="3a368-1172">文字列</span><span class="sxs-lookup"><span data-stu-id="3a368-1172">String</span></span>||<span data-ttu-id="3a368-p180">挿入されるデータです。データの長さは 1,000,000 文字以内です。1,000,000 文字を超えるデータが渡されると、 `ArgumentOutOfRange` の例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="3a368-p180">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="3a368-1176">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="3a368-1176">Object</span></span>| <span data-ttu-id="3a368-1177">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="3a368-1178">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="3a368-1178">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="3a368-1179">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="3a368-1179">Object</span></span>| <span data-ttu-id="3a368-1180">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="3a368-1181">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="3a368-1181">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="3a368-1182">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="3a368-1182">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="3a368-1183">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="3a368-1183">&lt;optional&gt;</span></span>|<span data-ttu-id="3a368-p181">`text` の場合、Outlook Web アプリ と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-p181">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="3a368-p182">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-p182">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="3a368-1188">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="3a368-1188">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="3a368-1189">関数</span><span class="sxs-lookup"><span data-stu-id="3a368-1189">function</span></span>||<span data-ttu-id="3a368-1190">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="3a368-1190">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3a368-1191">要件</span><span class="sxs-lookup"><span data-stu-id="3a368-1191">Requirements</span></span>

|<span data-ttu-id="3a368-1192">必要条件</span><span class="sxs-lookup"><span data-stu-id="3a368-1192">Requirement</span></span>| <span data-ttu-id="3a368-1193">値</span><span class="sxs-lookup"><span data-stu-id="3a368-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="3a368-1194">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3a368-1194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3a368-1195">1.2</span><span class="sxs-lookup"><span data-stu-id="3a368-1195">1.2</span></span>|
|[<span data-ttu-id="3a368-1196">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="3a368-1196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3a368-1197">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3a368-1197">ReadWriteItem</span></span>|
|[<span data-ttu-id="3a368-1198">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3a368-1198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3a368-1199">新規作成</span><span class="sxs-lookup"><span data-stu-id="3a368-1199">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="3a368-1200">例</span><span class="sxs-lookup"><span data-stu-id="3a368-1200">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```