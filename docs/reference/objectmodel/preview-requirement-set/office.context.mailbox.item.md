
# <a name="item"></a><span data-ttu-id="8639f-101">項目</span><span class="sxs-lookup"><span data-stu-id="8639f-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="8639f-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="8639f-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="8639f-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="8639f-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-105">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-105">Requirements</span></span>

|<span data-ttu-id="8639f-106">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-106">Requirement</span></span>|<span data-ttu-id="8639f-107">値</span><span class="sxs-lookup"><span data-stu-id="8639f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-108">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-109">1.0</span></span>|
|[<span data-ttu-id="8639f-110">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="8639f-111">Restricted</span></span>|
|[<span data-ttu-id="8639f-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-113">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8639f-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8639f-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-114">Members and methods</span></span>

| <span data-ttu-id="8639f-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-115">Member</span></span> | <span data-ttu-id="8639f-116">型</span><span class="sxs-lookup"><span data-stu-id="8639f-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8639f-117">attachments</span><span class="sxs-lookup"><span data-stu-id="8639f-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="8639f-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-118">Member</span></span> |
| [<span data-ttu-id="8639f-119">bcc</span><span class="sxs-lookup"><span data-stu-id="8639f-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8639f-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-120">Member</span></span> |
| [<span data-ttu-id="8639f-121">body</span><span class="sxs-lookup"><span data-stu-id="8639f-121">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="8639f-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-122">Member</span></span> |
| [<span data-ttu-id="8639f-123">cc</span><span class="sxs-lookup"><span data-stu-id="8639f-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8639f-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-124">Member</span></span> |
| [<span data-ttu-id="8639f-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="8639f-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="8639f-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-126">Member</span></span> |
| [<span data-ttu-id="8639f-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="8639f-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="8639f-128">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-128">Member</span></span> |
| [<span data-ttu-id="8639f-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="8639f-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="8639f-130">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-130">Member</span></span> |
| [<span data-ttu-id="8639f-131">end</span><span class="sxs-lookup"><span data-stu-id="8639f-131">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="8639f-132">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-132">Member</span></span> |
| [<span data-ttu-id="8639f-133">from</span><span class="sxs-lookup"><span data-stu-id="8639f-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="8639f-134">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-134">Member</span></span> |
| [<span data-ttu-id="8639f-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="8639f-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="8639f-136">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-136">Member</span></span> |
| [<span data-ttu-id="8639f-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="8639f-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="8639f-138">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-138">Member</span></span> |
| [<span data-ttu-id="8639f-139">itemId</span><span class="sxs-lookup"><span data-stu-id="8639f-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="8639f-140">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-140">Member</span></span> |
| [<span data-ttu-id="8639f-141">itemType</span><span class="sxs-lookup"><span data-stu-id="8639f-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="8639f-142">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-142">Member</span></span> |
| [<span data-ttu-id="8639f-143">location</span><span class="sxs-lookup"><span data-stu-id="8639f-143">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="8639f-144">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-144">Member</span></span> |
| [<span data-ttu-id="8639f-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="8639f-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="8639f-146">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-146">Member</span></span> |
| [<span data-ttu-id="8639f-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="8639f-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="8639f-148">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-148">Member</span></span> |
| [<span data-ttu-id="8639f-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="8639f-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8639f-150">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-150">Member</span></span> |
| [<span data-ttu-id="8639f-151">主催者</span><span class="sxs-lookup"><span data-stu-id="8639f-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="8639f-152">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-152">Member</span></span> |
| [<span data-ttu-id="8639f-153">パターン</span><span class="sxs-lookup"><span data-stu-id="8639f-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="8639f-154">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-154">Member</span></span> |
| [<span data-ttu-id="8639f-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="8639f-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8639f-156">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-156">Member</span></span> |
| [<span data-ttu-id="8639f-157">送り主</span><span class="sxs-lookup"><span data-stu-id="8639f-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="8639f-158">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-158">Member</span></span> |
| [<span data-ttu-id="8639f-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="8639f-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="8639f-160">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-160">Member</span></span> |
| [<span data-ttu-id="8639f-161">開始</span><span class="sxs-lookup"><span data-stu-id="8639f-161">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="8639f-162">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-162">Member</span></span> |
| [<span data-ttu-id="8639f-163">件名</span><span class="sxs-lookup"><span data-stu-id="8639f-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="8639f-164">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-164">Member</span></span> |
| [<span data-ttu-id="8639f-165">宛先</span><span class="sxs-lookup"><span data-stu-id="8639f-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8639f-166">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-166">Member</span></span> |
| [<span data-ttu-id="8639f-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8639f-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="8639f-168">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-168">Method</span></span> |
| [<span data-ttu-id="8639f-169">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="8639f-169">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="8639f-170">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-170">Method</span></span> |
| [<span data-ttu-id="8639f-171">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="8639f-171">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="8639f-172">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-172">Method</span></span> |
| [<span data-ttu-id="8639f-173">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8639f-173">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="8639f-174">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-174">Method</span></span> |
| [<span data-ttu-id="8639f-175">終了</span><span class="sxs-lookup"><span data-stu-id="8639f-175">close</span></span>](#close) | <span data-ttu-id="8639f-176">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-176">Method</span></span> |
| [<span data-ttu-id="8639f-177">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="8639f-177">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="8639f-178">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-178">Method</span></span> |
| [<span data-ttu-id="8639f-179">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="8639f-179">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="8639f-180">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-180">Method</span></span> |
| [<span data-ttu-id="8639f-181">getEntities</span><span class="sxs-lookup"><span data-stu-id="8639f-181">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="8639f-182">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-182">Method</span></span> |
| [<span data-ttu-id="8639f-183">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="8639f-183">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="8639f-184">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-184">Method</span></span> |
| [<span data-ttu-id="8639f-185">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="8639f-185">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="8639f-186">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-186">Method</span></span> |
| [<span data-ttu-id="8639f-187">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="8639f-187">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="8639f-188">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-188">Method</span></span> |
| [<span data-ttu-id="8639f-189">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="8639f-189">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="8639f-190">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-190">Method</span></span> |
| [<span data-ttu-id="8639f-191">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="8639f-191">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="8639f-192">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-192">Method</span></span> |
| [<span data-ttu-id="8639f-193">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="8639f-193">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="8639f-194">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-194">Method</span></span> |
| [<span data-ttu-id="8639f-195">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="8639f-195">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="8639f-196">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-196">Method</span></span> |
| [<span data-ttu-id="8639f-197">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="8639f-197">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="8639f-198">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-198">Method</span></span> |
| [<span data-ttu-id="8639f-199">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="8639f-199">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="8639f-200">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-200">Method</span></span> |
| [<span data-ttu-id="8639f-201">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="8639f-201">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="8639f-202">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-202">Method</span></span> |
| [<span data-ttu-id="8639f-203">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8639f-203">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="8639f-204">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-204">Method</span></span> |
| [<span data-ttu-id="8639f-205">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="8639f-205">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="8639f-206">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-206">Method</span></span> |
| [<span data-ttu-id="8639f-207">saveAsync</span><span class="sxs-lookup"><span data-stu-id="8639f-207">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="8639f-208">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-208">Method</span></span> |
| [<span data-ttu-id="8639f-209">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="8639f-209">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="8639f-210">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-210">Method</span></span> |

### <a name="example"></a><span data-ttu-id="8639f-211">例</span><span class="sxs-lookup"><span data-stu-id="8639f-211">Example</span></span>

<span data-ttu-id="8639f-212">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8639f-212">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="8639f-213">メンバー</span><span class="sxs-lookup"><span data-stu-id="8639f-213">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="8639f-214">添付ファイル：配列.<[添付ファイルの詳細 ](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="8639f-214">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="8639f-p102">項目の添付ファイルの配列を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="8639f-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8639f-217">潜在的なセキュリティ問題により特定の種類のファイルは、Outlookでブロックされ、したがって戻ってきません。</span><span class="sxs-lookup"><span data-stu-id="8639f-217">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="8639f-218">詳細については、[「Outlook でブロックされた添付ファイル」](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)をご覧下さい。</span><span class="sxs-lookup"><span data-stu-id="8639f-218">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-219">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-219">Type:</span></span>

*   <span data-ttu-id="8639f-220">配列.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="8639f-220">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-221">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-221">Requirements</span></span>

|<span data-ttu-id="8639f-222">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-222">Requirement</span></span>|<span data-ttu-id="8639f-223">値</span><span class="sxs-lookup"><span data-stu-id="8639f-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-224">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-224">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-225">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-225">1.0</span></span>|
|[<span data-ttu-id="8639f-226">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-227">ReadItem</span></span>|
|[<span data-ttu-id="8639f-228">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-229">読み取り</span><span class="sxs-lookup"><span data-stu-id="8639f-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-230">例</span><span class="sxs-lookup"><span data-stu-id="8639f-230">Example</span></span>

<span data-ttu-id="8639f-231">次のコードでは、現在の項目にあるすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="8639f-231">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8639f-232">bcc:[受信者](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8639f-232">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8639f-233">メッセージの BCC (ブラインド カーボン コピー) 列上の 受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="8639f-233">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="8639f-234">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8639f-234">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-235">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-235">Type:</span></span>

*   [<span data-ttu-id="8639f-236">受信者</span><span class="sxs-lookup"><span data-stu-id="8639f-236">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="8639f-237">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-237">Requirements</span></span>

|<span data-ttu-id="8639f-238">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-238">Requirement</span></span>|<span data-ttu-id="8639f-239">値</span><span class="sxs-lookup"><span data-stu-id="8639f-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-240">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-241">1.1</span><span class="sxs-lookup"><span data-stu-id="8639f-241">1.1</span></span>|
|[<span data-ttu-id="8639f-242">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-243">ReadItem</span></span>|
|[<span data-ttu-id="8639f-244">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-245">新規作成</span><span class="sxs-lookup"><span data-stu-id="8639f-245">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-246">例</span><span class="sxs-lookup"><span data-stu-id="8639f-246">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="8639f-247">本文:[本文](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="8639f-247">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="8639f-248">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="8639f-248">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-249">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-249">Type:</span></span>

*   [<span data-ttu-id="8639f-250">本文</span><span class="sxs-lookup"><span data-stu-id="8639f-250">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="8639f-251">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-251">Requirements</span></span>

|<span data-ttu-id="8639f-252">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-252">Requirement</span></span>|<span data-ttu-id="8639f-253">値</span><span class="sxs-lookup"><span data-stu-id="8639f-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-254">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-254">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-255">1.1</span><span class="sxs-lookup"><span data-stu-id="8639f-255">1.1</span></span>|
|[<span data-ttu-id="8639f-256">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-256">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-257">ReadItem</span></span>|
|[<span data-ttu-id="8639f-258">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-258">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-259">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8639f-259">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8639f-260">cc: 配列。 <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[受信者](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8639f-260">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8639f-261">メッセージの CC (カーボン コピー) 受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8639f-261">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="8639f-262">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="8639f-262">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8639f-263">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8639f-263">Read mode</span></span>

<span data-ttu-id="8639f-p106">`cc`プロパティは、メッセージの**CC**列にある各受信者一覧の`EmailAddressDetails`オブジェクトを含む配列を返します。コレクションは最大100個のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8639f-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8639f-266">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8639f-266">Compose mode</span></span>

<span data-ttu-id="8639f-267">`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-267">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-268">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-268">Type:</span></span>

*   <span data-ttu-id="8639f-269">配列 。<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> |[受信者](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8639f-269">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-270">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-270">Requirements</span></span>

|<span data-ttu-id="8639f-271">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-271">Requirement</span></span>|<span data-ttu-id="8639f-272">値</span><span class="sxs-lookup"><span data-stu-id="8639f-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-273">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-274">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-274">1.0</span></span>|
|[<span data-ttu-id="8639f-275">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-276">ReadItem</span></span>|
|[<span data-ttu-id="8639f-277">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-278">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8639f-278">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-279">例</span><span class="sxs-lookup"><span data-stu-id="8639f-279">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="8639f-280">（空白が可能）conversationId：文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-280">(nullable) conversationId :String</span></span>

<span data-ttu-id="8639f-281">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="8639f-281">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="8639f-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="8639f-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="8639f-p108">作成フォームの新しいアイテムに対してこのプロパティの Null を取得します。ユーザーが件名を設定し項目を保存する場合、`conversationId`プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-286">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-286">Type:</span></span>

*   <span data-ttu-id="8639f-287">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-287">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-288">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-288">Requirements</span></span>

|<span data-ttu-id="8639f-289">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-289">Requirement</span></span>|<span data-ttu-id="8639f-290">値</span><span class="sxs-lookup"><span data-stu-id="8639f-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-291">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-292">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-292">1.0</span></span>|
|[<span data-ttu-id="8639f-293">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-294">ReadItem</span></span>|
|[<span data-ttu-id="8639f-295">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-296">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8639f-296">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="8639f-297">dateTimeCreated: 日付</span><span class="sxs-lookup"><span data-stu-id="8639f-297">dateTimeCreated :Date</span></span>

<span data-ttu-id="8639f-p109">アイテムが作成された日時を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="8639f-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-300">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-300">Type:</span></span>

*   <span data-ttu-id="8639f-301">日付</span><span class="sxs-lookup"><span data-stu-id="8639f-301">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-302">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-302">Requirements</span></span>

|<span data-ttu-id="8639f-303">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-303">Requirement</span></span>|<span data-ttu-id="8639f-304">値</span><span class="sxs-lookup"><span data-stu-id="8639f-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-305">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-306">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-306">1.0</span></span>|
|[<span data-ttu-id="8639f-307">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-308">ReadItem</span></span>|
|[<span data-ttu-id="8639f-309">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-310">読み取り</span><span class="sxs-lookup"><span data-stu-id="8639f-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-311">例</span><span class="sxs-lookup"><span data-stu-id="8639f-311">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="8639f-312">dateTimeModified: 日付</span><span class="sxs-lookup"><span data-stu-id="8639f-312">dateTimeModified :Date</span></span>

<span data-ttu-id="8639f-p110">アイテムが最後に変更された日時を取得します。読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="8639f-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8639f-315">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8639f-315">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-316">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-316">Type:</span></span>

*   <span data-ttu-id="8639f-317">日付</span><span class="sxs-lookup"><span data-stu-id="8639f-317">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-318">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-318">Requirements</span></span>

|<span data-ttu-id="8639f-319">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-319">Requirement</span></span>|<span data-ttu-id="8639f-320">値</span><span class="sxs-lookup"><span data-stu-id="8639f-320">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-321">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-321">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-322">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-322">1.0</span></span>|
|[<span data-ttu-id="8639f-323">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-323">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-324">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-324">ReadItem</span></span>|
|[<span data-ttu-id="8639f-325">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-325">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-326">読み取り</span><span class="sxs-lookup"><span data-stu-id="8639f-326">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-327">例</span><span class="sxs-lookup"><span data-stu-id="8639f-327">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="8639f-328">end:日付|[時間](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="8639f-328">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="8639f-329">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8639f-329">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="8639f-p111">`end`プロパティは、協定世界時 (UTC) 形式の時刻値として表示されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="8639f-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8639f-332">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8639f-332">Read mode</span></span>

<span data-ttu-id="8639f-333">`end`プロパティは`Date`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-333">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8639f-334">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8639f-334">Compose mode</span></span>

<span data-ttu-id="8639f-335">`end`プロパティは`Time`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-335">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="8639f-336">[ `Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-)メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8639f-336">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-337">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-337">Type:</span></span>

*   <span data-ttu-id="8639f-338">日付| [時間](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="8639f-338">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-339">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-339">Requirements</span></span>

|<span data-ttu-id="8639f-340">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-340">Requirement</span></span>|<span data-ttu-id="8639f-341">値</span><span class="sxs-lookup"><span data-stu-id="8639f-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-342">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-342">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-343">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-343">1.0</span></span>|
|[<span data-ttu-id="8639f-344">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-344">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-345">ReadItem</span></span>|
|[<span data-ttu-id="8639f-346">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-346">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-347">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8639f-347">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-348">例</span><span class="sxs-lookup"><span data-stu-id="8639f-348">Example</span></span>

<span data-ttu-id="8639f-349">次の例では、`Time`オブジェクトの[`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-)メソッドを使用して、作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="8639f-349">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="8639f-350">:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[から](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="8639f-350">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="8639f-351">メッセージの送信者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="8639f-351">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="8639f-p112">メッセージが代理人から送信された場合を除き、`from` プロパティと[`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails)プロパティは同一人物を表します。その場合、`from` プロパティは委任者を、送信者プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="8639f-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8639f-354">`from`プロパティ内の`EmailAddressDetails`オブジェクトの`recipientType`プロパティは、`undefined`です。</span><span class="sxs-lookup"><span data-stu-id="8639f-354">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8639f-355">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8639f-355">Read mode</span></span>

<span data-ttu-id="8639f-356">`from`プロパティは`EmailAddressDetails`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-356">The `from` property returns an `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="8639f-357">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8639f-357">Compose mode</span></span>

<span data-ttu-id="8639f-358">`from`プロパティは送信者値を取得するためのメソッドを提供する`From`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-358">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8639f-359">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-359">Type:</span></span>

*   <span data-ttu-id="8639f-360">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [から](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="8639f-360">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-361">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-361">Requirements</span></span>

|<span data-ttu-id="8639f-362">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-362">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="8639f-363">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-364">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-364">1.0</span></span>|<span data-ttu-id="8639f-365">1.7</span><span class="sxs-lookup"><span data-stu-id="8639f-365">-17</span></span>|
|[<span data-ttu-id="8639f-366">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-366">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-367">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-367">ReadItem</span></span>|<span data-ttu-id="8639f-368">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8639f-368">ReadWriteItem</span></span>|
|[<span data-ttu-id="8639f-369">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-369">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-370">読み取り</span><span class="sxs-lookup"><span data-stu-id="8639f-370">Read</span></span>|<span data-ttu-id="8639f-371">新規作成</span><span class="sxs-lookup"><span data-stu-id="8639f-371">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="8639f-372">internetMessageId:文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-372">internetMessageId :String</span></span>

<span data-ttu-id="8639f-p113">電子メール メッセージのインターネット メッセージ 識別子を取得します。読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="8639f-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-375">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-375">Type:</span></span>

*   <span data-ttu-id="8639f-376">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-376">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-377">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-377">Requirements</span></span>

|<span data-ttu-id="8639f-378">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-378">Requirement</span></span>|<span data-ttu-id="8639f-379">値</span><span class="sxs-lookup"><span data-stu-id="8639f-379">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-380">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-380">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-381">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-381">1.0</span></span>|
|[<span data-ttu-id="8639f-382">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-382">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-383">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-383">ReadItem</span></span>|
|[<span data-ttu-id="8639f-384">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-384">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-385">読み取り</span><span class="sxs-lookup"><span data-stu-id="8639f-385">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-386">例</span><span class="sxs-lookup"><span data-stu-id="8639f-386">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="8639f-387">itemClass: 文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-387">itemClass :String</span></span>

<span data-ttu-id="8639f-p114">選択された項目の Exchange Web サービスの項目 クラスを取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="8639f-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="8639f-p115">`itemClass` プロパティは、選択したアイテムのメッセージ クラスを指定します。次は、メッセージまたは予定アイテムの既定のメッセージ クラスを示しています。</span><span class="sxs-lookup"><span data-stu-id="8639f-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="8639f-392">種類</span><span class="sxs-lookup"><span data-stu-id="8639f-392">Type</span></span>|<span data-ttu-id="8639f-393">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-393">Description</span></span>|<span data-ttu-id="8639f-394">項目のクラス</span><span class="sxs-lookup"><span data-stu-id="8639f-394">item class</span></span>|
|---|---|---|
|<span data-ttu-id="8639f-395">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="8639f-395">Appointment items</span></span>|<span data-ttu-id="8639f-396">これらは、アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurence` の予定表アイテムです。</span><span class="sxs-lookup"><span data-stu-id="8639f-396">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="8639f-397">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="8639f-397">Message items</span></span>|<span data-ttu-id="8639f-398">これには、基本のメッセージ クラス として `IPM.Schedule.Meeting`  を使用する、既定のメッセージ クラス `IPM.Note`  会議出席依頼、返信および取り消しを持つ電子メール メッセージが含まれます。</span><span class="sxs-lookup"><span data-stu-id="8639f-398">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="8639f-399">既定のメッセージ クラスを拡張したカスタム メッセージ クラス 、たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など)を作成できます。</span><span class="sxs-lookup"><span data-stu-id="8639f-399">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-400">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-400">Type:</span></span>

*   <span data-ttu-id="8639f-401">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-401">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-402">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-402">Requirements</span></span>

|<span data-ttu-id="8639f-403">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-403">Requirement</span></span>|<span data-ttu-id="8639f-404">値</span><span class="sxs-lookup"><span data-stu-id="8639f-404">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-405">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-405">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-406">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-406">1.0</span></span>|
|[<span data-ttu-id="8639f-407">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-407">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-408">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-408">ReadItem</span></span>|
|[<span data-ttu-id="8639f-409">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-409">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-410">読み取り</span><span class="sxs-lookup"><span data-stu-id="8639f-410">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-411">例</span><span class="sxs-lookup"><span data-stu-id="8639f-411">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="8639f-412">（空白が可能） itemId ：文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-412">(nullable) itemId :String</span></span>

<span data-ttu-id="8639f-p116">現在の項目の Exchange Web サービスのアイテム識別子を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="8639f-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8639f-415">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="8639f-415">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="8639f-416">`itemId` プロパティは、Outlook の Entry ID または Outlook の REST API によって使用される ID と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="8639f-416">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="8639f-417">この値を使用して REST API の呼び出しを行う前に [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8639f-417">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="8639f-418">詳細については、 [Outlook アドインから Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8639f-418">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="8639f-p118">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-421">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-421">Type:</span></span>

*   <span data-ttu-id="8639f-422">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-422">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-423">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-423">Requirements</span></span>

|<span data-ttu-id="8639f-424">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-424">Requirement</span></span>|<span data-ttu-id="8639f-425">値</span><span class="sxs-lookup"><span data-stu-id="8639f-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-426">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-427">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-427">1.0</span></span>|
|[<span data-ttu-id="8639f-428">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-428">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-429">ReadItem</span></span>|
|[<span data-ttu-id="8639f-430">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-430">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-431">読み取り</span><span class="sxs-lookup"><span data-stu-id="8639f-431">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-432">例</span><span class="sxs-lookup"><span data-stu-id="8639f-432">Example</span></span>

<span data-ttu-id="8639f-p119">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="8639f-435">itemType:[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="8639f-435">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="8639f-436">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="8639f-436">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="8639f-437">`itemType`プロパティは、`ItemType`列挙値の 1 つを返します。これは`item`オブジェクト インスタンスがメッセージまたは予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="8639f-437">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-438">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-438">Type:</span></span>

*   [<span data-ttu-id="8639f-439">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="8639f-439">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="8639f-440">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-440">Requirements</span></span>

|<span data-ttu-id="8639f-441">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-441">Requirement</span></span>|<span data-ttu-id="8639f-442">値</span><span class="sxs-lookup"><span data-stu-id="8639f-442">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-443">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-443">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-444">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-444">1.0</span></span>|
|[<span data-ttu-id="8639f-445">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-445">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-446">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-446">ReadItem</span></span>|
|[<span data-ttu-id="8639f-447">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-447">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-448">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8639f-448">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-449">例</span><span class="sxs-lookup"><span data-stu-id="8639f-449">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="8639f-450">位置: 文字列|[](/javascript/api/outlook/office.location)位置</span><span class="sxs-lookup"><span data-stu-id="8639f-450">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="8639f-451">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8639f-451">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8639f-452">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8639f-452">Read mode</span></span>

<span data-ttu-id="8639f-453">`location` プロパティは、予定の場所を含む文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-453">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8639f-454">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8639f-454">Compose mode</span></span>

<span data-ttu-id="8639f-455">`location` プロパティは、予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-455">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-456">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-456">Type:</span></span>

*   <span data-ttu-id="8639f-457">文字列 | [場所](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="8639f-457">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-458">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-458">Requirements</span></span>

|<span data-ttu-id="8639f-459">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-459">Requirement</span></span>|<span data-ttu-id="8639f-460">値</span><span class="sxs-lookup"><span data-stu-id="8639f-460">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-461">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-461">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-462">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-462">1.0</span></span>|
|[<span data-ttu-id="8639f-463">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-463">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-464">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-464">ReadItem</span></span>|
|[<span data-ttu-id="8639f-465">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-465">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-466">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8639f-466">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-467">例</span><span class="sxs-lookup"><span data-stu-id="8639f-467">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="8639f-468">normalizedSubject: 文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-468">normalizedSubject :String</span></span>

<span data-ttu-id="8639f-p120">すべてのプレフィックス (`RE:` や `FWD:` を含む) が削除された項目の件名を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="8639f-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="8639f-p121">normalizedSubject プロパティは、電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたアイテムの件名を取得します。プレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="8639f-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-473">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-473">Type:</span></span>

*   <span data-ttu-id="8639f-474">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-474">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-475">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-475">Requirements</span></span>

|<span data-ttu-id="8639f-476">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-476">Requirement</span></span>|<span data-ttu-id="8639f-477">値</span><span class="sxs-lookup"><span data-stu-id="8639f-477">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-478">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-478">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-479">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-479">1.0</span></span>|
|[<span data-ttu-id="8639f-480">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-480">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-481">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-481">ReadItem</span></span>|
|[<span data-ttu-id="8639f-482">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-482">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-483">読み取り</span><span class="sxs-lookup"><span data-stu-id="8639f-483">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-484">例</span><span class="sxs-lookup"><span data-stu-id="8639f-484">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="8639f-485">notificationMessages:[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="8639f-485">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="8639f-486">項目の通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="8639f-486">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-487">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-487">Type:</span></span>

*   [<span data-ttu-id="8639f-488">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="8639f-488">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="8639f-489">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-489">Requirements</span></span>

|<span data-ttu-id="8639f-490">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-490">Requirement</span></span>|<span data-ttu-id="8639f-491">値</span><span class="sxs-lookup"><span data-stu-id="8639f-491">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-492">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-492">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-493">1.3</span><span class="sxs-lookup"><span data-stu-id="8639f-493">1.3</span></span>|
|[<span data-ttu-id="8639f-494">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-494">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-495">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-495">ReadItem</span></span>|
|[<span data-ttu-id="8639f-496">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-496">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-497">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8639f-497">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8639f-498">optionalAttendees: 配列.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[受信者](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8639f-498">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8639f-499">イベントの任意の出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8639f-499">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="8639f-500">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="8639f-500">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8639f-501">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8639f-501">Read mode</span></span>

<span data-ttu-id="8639f-502">`optionalAttendees`プロパティは、会議への各任意出席者の`EmailAddressDetails`オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-502">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8639f-503">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8639f-503">Compose mode</span></span>

<span data-ttu-id="8639f-504">`optionalAttendees`プロパティは会議への任意出席者を取得または設定するためのメソッドを提供する`Recipients`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-504">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-505">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-505">Type:</span></span>

*   <span data-ttu-id="8639f-506">配列 。<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> |[受信者](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8639f-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-507">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-507">Requirements</span></span>

|<span data-ttu-id="8639f-508">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-508">Requirement</span></span>|<span data-ttu-id="8639f-509">値</span><span class="sxs-lookup"><span data-stu-id="8639f-509">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-510">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-510">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-511">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-511">1.0</span></span>|
|[<span data-ttu-id="8639f-512">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-512">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-513">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-513">ReadItem</span></span>|
|[<span data-ttu-id="8639f-514">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-514">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-515">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8639f-515">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-516">例</span><span class="sxs-lookup"><span data-stu-id="8639f-516">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="8639f-517">開催際者:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="8639f-517">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="8639f-518">指定の会議の開催者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="8639f-518">Gets the email address of the meeting organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8639f-519">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8639f-519">Read mode</span></span>

<span data-ttu-id="8639f-520">`organizer`プロパティは、会議の開催者を表す[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-520">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8639f-521">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8639f-521">Compose mode</span></span>

<span data-ttu-id="8639f-522">`organizer`プロパティが開催者の値を取得するメソッドを提供する[Organizer](/javascript/api/outlook/office.organizer)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-522">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-523">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-523">Type:</span></span>

*   <span data-ttu-id="8639f-524">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="8639f-524">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-525">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-525">Requirements</span></span>

|<span data-ttu-id="8639f-526">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-526">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="8639f-527">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-527">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-528">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-528">1.0</span></span>|<span data-ttu-id="8639f-529">1.7</span><span class="sxs-lookup"><span data-stu-id="8639f-529">-17</span></span>|
|[<span data-ttu-id="8639f-530">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-530">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-531">ReadItem</span></span>|<span data-ttu-id="8639f-532">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8639f-532">ReadWriteItem</span></span>|
|[<span data-ttu-id="8639f-533">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-533">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-534">読み取り</span><span class="sxs-lookup"><span data-stu-id="8639f-534">Read</span></span>|<span data-ttu-id="8639f-535">新規作成</span><span class="sxs-lookup"><span data-stu-id="8639f-535">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-536">例</span><span class="sxs-lookup"><span data-stu-id="8639f-536">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="8639f-537">(Null 許容) 定期的: [Recurrence](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="8639f-537">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="8639f-538">予定の定期的なパターンを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8639f-538">Gets or sets the location of an appointment.</span></span> <span data-ttu-id="8639f-539">会議出席依頼の定期的なパターンを取得します。</span><span class="sxs-lookup"><span data-stu-id="8639f-539">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="8639f-540">予定表アイテムの読み込みモードおよび作成モードです。</span><span class="sxs-lookup"><span data-stu-id="8639f-540">Read and compose modes for appointment items.</span></span> <span data-ttu-id="8639f-541">会議出席依頼アイテムの読み取りモードです。</span><span class="sxs-lookup"><span data-stu-id="8639f-541">Read mode for meeting request items.</span></span>

<span data-ttu-id="8639f-542">`recurrence`プロパティは、アイテムが系列または系列のインスタンスである場合に定期的な予定または会議出席依頼に[定期的な](/javascript/api/outlook/office.recurrence)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-542">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="8639f-543">`null` 単独の予定および単独の予定の会議出席依頼に返されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-543">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="8639f-544">`undefined` 会議出席依頼ではないメッセージに返されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-544">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="8639f-545">注: 会議出席依頼は、IPM.Schedule.Meeting.Request の`itemClass`値を含んでいます。</span><span class="sxs-lookup"><span data-stu-id="8639f-545">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="8639f-546">注: 定期的なオブジェクトが`null`である場合、これは、オブジェクトが 単独の予定または会議出席依頼、単独の予定および系列の一部ではないことを示します。</span><span class="sxs-lookup"><span data-stu-id="8639f-546">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-547">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-547">Type:</span></span>

* [<span data-ttu-id="8639f-548">パターン</span><span class="sxs-lookup"><span data-stu-id="8639f-548">recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="8639f-549">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-549">Requirement</span></span>|<span data-ttu-id="8639f-550">値</span><span class="sxs-lookup"><span data-stu-id="8639f-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-551">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-552">1.7</span><span class="sxs-lookup"><span data-stu-id="8639f-552">-17</span></span>|
|[<span data-ttu-id="8639f-553">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-554">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-554">ReadItem</span></span>|
|[<span data-ttu-id="8639f-555">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-556">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8639f-556">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8639f-557">requiredAttendees: 配列.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8639f-557">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8639f-558">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8639f-558">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="8639f-559">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="8639f-559">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8639f-560">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8639f-560">Read mode</span></span>

<span data-ttu-id="8639f-561">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-561">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8639f-562">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8639f-562">Compose mode</span></span>

<span data-ttu-id="8639f-563">`requiredAttendees` プロパティは会議への必須出席者を取得または設定するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-563">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-564">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-564">Type:</span></span>

*   <span data-ttu-id="8639f-565">配列 。<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> |[受信者](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8639f-565">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-566">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-566">Requirements</span></span>

|<span data-ttu-id="8639f-567">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-567">Requirement</span></span>|<span data-ttu-id="8639f-568">値</span><span class="sxs-lookup"><span data-stu-id="8639f-568">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-569">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-569">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-570">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-570">1.0</span></span>|
|[<span data-ttu-id="8639f-571">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-571">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-572">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-572">ReadItem</span></span>|
|[<span data-ttu-id="8639f-573">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-573">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-574">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8639f-574">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-575">例</span><span class="sxs-lookup"><span data-stu-id="8639f-575">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="8639f-576">送信者:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="8639f-576">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="8639f-p126">電子メール送信者のメールアドレスを取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="8639f-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="8639f-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) プロパティと `sender` プロパティは同一人物を表します。その場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="8639f-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8639f-581">`sender`プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType`プロパティは、`undefined`です。</span><span class="sxs-lookup"><span data-stu-id="8639f-581">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-582">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-582">Type:</span></span>

*   [<span data-ttu-id="8639f-583">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8639f-583">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="8639f-584">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-584">Requirements</span></span>

|<span data-ttu-id="8639f-585">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-585">Requirement</span></span>|<span data-ttu-id="8639f-586">値</span><span class="sxs-lookup"><span data-stu-id="8639f-586">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-587">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-587">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-588">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-588">1.0</span></span>|
|[<span data-ttu-id="8639f-589">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-589">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-590">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-590">ReadItem</span></span>|
|[<span data-ttu-id="8639f-591">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-591">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-592">読み取り</span><span class="sxs-lookup"><span data-stu-id="8639f-592">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-593">例</span><span class="sxs-lookup"><span data-stu-id="8639f-593">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="8639f-594">(Null 許容) seriesId: 文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-594">(nullable) seriesId :String</span></span>

<span data-ttu-id="8639f-595">インスタンスが属する系列の ID を取得します。</span><span class="sxs-lookup"><span data-stu-id="8639f-595">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="8639f-596">OWA と Outlook で、 `seriesId`は、この項目が属する親 (系列) アイテムの Exchange Web サービス (EWS) の ID を返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-596">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="8639f-597">しかし、IOS および Android で、`seriesId`、親項目の REST ID を返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-597">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="8639f-598">`seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="8639f-598">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="8639f-599">`seriesId`プロパティは Outlook の REST API で使用される Outlook ID と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="8639f-599">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="8639f-600">この値を使用して REST API の呼び出しを行う前に [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8639f-600">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="8639f-601">詳細については、[「Outlook アドインから Outlook REST API の使用」](https://docs.microsoft.com/outlook/add-ins/use-rest-api)をご覧下さい。</span><span class="sxs-lookup"><span data-stu-id="8639f-601">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="8639f-602">`seriesId`プロパティは、単一の予定、系列のアイテム、または会議出席依頼などの親アイテムを持たないには`null` を返し、会議出席依頼ではないその他のアイテムには`undefined`を返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-602">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-603">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-603">Type:</span></span>

* <span data-ttu-id="8639f-604">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-604">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-605">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-605">Requirements</span></span>

|<span data-ttu-id="8639f-606">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-606">Requirement</span></span>|<span data-ttu-id="8639f-607">値</span><span class="sxs-lookup"><span data-stu-id="8639f-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-608">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-609">1.7</span><span class="sxs-lookup"><span data-stu-id="8639f-609">-17</span></span>|
|[<span data-ttu-id="8639f-610">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-610">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-611">ReadItem</span></span>|
|[<span data-ttu-id="8639f-612">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-612">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-613">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8639f-613">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-614">例</span><span class="sxs-lookup"><span data-stu-id="8639f-614">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="8639f-615">開始: 日付 | [ 時間](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="8639f-615">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="8639f-616">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8639f-616">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="8639f-p130">`start`プロパティは、協定世界時 (UTC) 形式の時刻値として表示されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime)メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="8639f-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8639f-619">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8639f-619">Read mode</span></span>

<span data-ttu-id="8639f-620">`start`プロパティは`Date`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-620">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8639f-621">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8639f-621">Compose mode</span></span>

<span data-ttu-id="8639f-622">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-622">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="8639f-623">[ `Time.setAsync` ](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8639f-623">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-624">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-624">Type:</span></span>

*   <span data-ttu-id="8639f-625">日付| [時間](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="8639f-625">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-626">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-626">Requirements</span></span>

|<span data-ttu-id="8639f-627">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-627">Requirement</span></span>|<span data-ttu-id="8639f-628">値</span><span class="sxs-lookup"><span data-stu-id="8639f-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-629">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-629">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-630">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-630">1.0</span></span>|
|[<span data-ttu-id="8639f-631">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-631">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-632">ReadItem</span></span>|
|[<span data-ttu-id="8639f-633">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-633">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-634">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8639f-634">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-635">例</span><span class="sxs-lookup"><span data-stu-id="8639f-635">Example</span></span>

<span data-ttu-id="8639f-636">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="8639f-636">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="8639f-637">件名: 文字列 | [件名](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="8639f-637">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="8639f-638">アイテムの件名フィールドに表示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8639f-638">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="8639f-639">`subject`プロパティは、電子メールサーバーから送信された項目の全件名を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8639f-639">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8639f-640">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8639f-640">Read mode</span></span>

<span data-ttu-id="8639f-p131">`subject`プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string)プロパティを使用して、行間にある`RE:`や`FW:`のなどのすべてのプレフィックスを削除した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="8639f-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="8639f-643">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8639f-643">Compose mode</span></span>

<span data-ttu-id="8639f-644">`subject`プロパティは、件名を取得または設定するためのメソッドを提供する`Subject`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-644">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8639f-645">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-645">Type:</span></span>

*   <span data-ttu-id="8639f-646">文字列 | [件名](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="8639f-646">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-647">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-647">Requirements</span></span>

|<span data-ttu-id="8639f-648">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-648">Requirement</span></span>|<span data-ttu-id="8639f-649">値</span><span class="sxs-lookup"><span data-stu-id="8639f-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-650">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-651">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-651">1.0</span></span>|
|[<span data-ttu-id="8639f-652">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-653">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-653">ReadItem</span></span>|
|[<span data-ttu-id="8639f-654">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-655">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8639f-655">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8639f-656">to: 配列。[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[受信者](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8639f-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8639f-657">メッセージの **宛先**列にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8639f-657">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="8639f-658">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="8639f-658">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8639f-659">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8639f-659">Read mode</span></span>

<span data-ttu-id="8639f-p133">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8639f-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8639f-662">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8639f-662">Compose mode</span></span>

<span data-ttu-id="8639f-663">`to` プロパティは、メッセージの **To** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-663">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="8639f-664">種類:</span><span class="sxs-lookup"><span data-stu-id="8639f-664">Type:</span></span>

*   <span data-ttu-id="8639f-665">配列 。<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> |[受信者](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8639f-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-666">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-666">Requirements</span></span>

|<span data-ttu-id="8639f-667">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-667">Requirement</span></span>|<span data-ttu-id="8639f-668">値</span><span class="sxs-lookup"><span data-stu-id="8639f-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-669">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-670">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-670">1.0</span></span>|
|[<span data-ttu-id="8639f-671">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-671">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-672">ReadItem</span></span>|
|[<span data-ttu-id="8639f-673">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-673">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-674">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8639f-674">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-675">例</span><span class="sxs-lookup"><span data-stu-id="8639f-675">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="8639f-676">メソッド</span><span class="sxs-lookup"><span data-stu-id="8639f-676">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="8639f-677">addFileAttachmentAsync(uri, attachmentName, [オプション], [コールバック])</span><span class="sxs-lookup"><span data-stu-id="8639f-677">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8639f-678">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="8639f-678">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="8639f-679">`addFileAttachmentAsync`メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内の項目に添付します。</span><span class="sxs-lookup"><span data-stu-id="8639f-679">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="8639f-680">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="8639f-680">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8639f-681">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="8639f-681">Parameters:</span></span>
|<span data-ttu-id="8639f-682">名前</span><span class="sxs-lookup"><span data-stu-id="8639f-682">Name</span></span>|<span data-ttu-id="8639f-683">種類</span><span class="sxs-lookup"><span data-stu-id="8639f-683">Type</span></span>|<span data-ttu-id="8639f-684">属性</span><span class="sxs-lookup"><span data-stu-id="8639f-684">Attributes</span></span>|<span data-ttu-id="8639f-685">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-685">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="8639f-686">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-686">String</span></span>||<span data-ttu-id="8639f-p134">メッセージまたは予定に添付するファイルの場所を示す URIです。 2048 文字以内で入力してください。</span><span class="sxs-lookup"><span data-stu-id="8639f-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="8639f-689">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-689">String</span></span>||<span data-ttu-id="8639f-p135">アップロード中に表示される添付ファイルがそのファイルの名前です。 255 文字以内で入力してください。</span><span class="sxs-lookup"><span data-stu-id="8639f-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="8639f-692">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-692">Object</span></span>|<span data-ttu-id="8639f-693">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-693">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-694">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="8639f-694">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8639f-695">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-695">Object</span></span>|<span data-ttu-id="8639f-696">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-696">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-697">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8639f-697">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="8639f-698">ブール値</span><span class="sxs-lookup"><span data-stu-id="8639f-698">Boolean</span></span>|<span data-ttu-id="8639f-699">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-699">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-700">`true`の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="8639f-700">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="8639f-701">関数</span><span class="sxs-lookup"><span data-stu-id="8639f-701">function</span></span>|<span data-ttu-id="8639f-702">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-702">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-703">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="8639f-703">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8639f-704">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-704">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8639f-705">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="8639f-705">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8639f-706">エラー</span><span class="sxs-lookup"><span data-stu-id="8639f-706">Errors</span></span>

|<span data-ttu-id="8639f-707">エラー コード</span><span class="sxs-lookup"><span data-stu-id="8639f-707">Error code</span></span>|<span data-ttu-id="8639f-708">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-708">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="8639f-709">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="8639f-709">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="8639f-710">許可されていない拡張子付きの添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="8639f-710">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="8639f-711">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="8639f-711">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8639f-712">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-712">Requirements</span></span>

|<span data-ttu-id="8639f-713">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-713">Requirement</span></span>|<span data-ttu-id="8639f-714">値</span><span class="sxs-lookup"><span data-stu-id="8639f-714">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-715">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-715">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-716">1.1</span><span class="sxs-lookup"><span data-stu-id="8639f-716">1.1</span></span>|
|[<span data-ttu-id="8639f-717">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-717">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-718">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8639f-718">ReadWriteItem</span></span>|
|[<span data-ttu-id="8639f-719">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-719">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-720">新規作成</span><span class="sxs-lookup"><span data-stu-id="8639f-720">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8639f-721">例</span><span class="sxs-lookup"><span data-stu-id="8639f-721">Examples</span></span>

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

<span data-ttu-id="8639f-722">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="8639f-722">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="8639f-723">addFileAttachmentFromBase64Async (base64File、attachmentName、[オプション]、[コールバック])</span><span class="sxs-lookup"><span data-stu-id="8639f-723">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8639f-724">メッセージまたは予定を添付ファイルとしてエンコード base64 からファイルを追加します。</span><span class="sxs-lookup"><span data-stu-id="8639f-724">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="8639f-725"> `addFileAttachmentFromBase64Async` メソッドは、base64 エンコーディングからファイルをアップロードし、作成フォーム内の項目にアタッチします。</span><span class="sxs-lookup"><span data-stu-id="8639f-725">The `addFileAttachmentFromBase64Async` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span> <span data-ttu-id="8639f-726">このメソッドは、AsyncResult.value オブジェクトの添付ファイルの識別子を返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-726">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="8639f-727">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="8639f-727">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8639f-728">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="8639f-728">Parameters:</span></span>
|<span data-ttu-id="8639f-729">名前</span><span class="sxs-lookup"><span data-stu-id="8639f-729">Name</span></span>|<span data-ttu-id="8639f-730">種類</span><span class="sxs-lookup"><span data-stu-id="8639f-730">Type</span></span>|<span data-ttu-id="8639f-731">属性</span><span class="sxs-lookup"><span data-stu-id="8639f-731">Attributes</span></span>|<span data-ttu-id="8639f-732">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-732">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="8639f-733">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-733">String</span></span>||<span data-ttu-id="8639f-734">電子メールまたはイベントに追加するイメージやファイルのコンテンツが base64 にエンコードされます。</span><span class="sxs-lookup"><span data-stu-id="8639f-734">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="8639f-735">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-735">String</span></span>||<span data-ttu-id="8639f-p137">アップロード中に表示される添付ファイルがそのファイルの名前です。 255 文字以内で入力してください。</span><span class="sxs-lookup"><span data-stu-id="8639f-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="8639f-738">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-738">Object</span></span>|<span data-ttu-id="8639f-739">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-739">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-740">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="8639f-740">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8639f-741">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-741">Object</span></span>|<span data-ttu-id="8639f-742">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-742">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-743">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8639f-743">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="8639f-744">ブール値</span><span class="sxs-lookup"><span data-stu-id="8639f-744">Boolean</span></span>|<span data-ttu-id="8639f-745">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-745">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-746">`true`の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="8639f-746">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="8639f-747">関数</span><span class="sxs-lookup"><span data-stu-id="8639f-747">function</span></span>|<span data-ttu-id="8639f-748">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-748">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-749">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="8639f-749">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8639f-750">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-750">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8639f-751">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="8639f-751">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8639f-752">エラー</span><span class="sxs-lookup"><span data-stu-id="8639f-752">Errors</span></span>

|<span data-ttu-id="8639f-753">エラー コード</span><span class="sxs-lookup"><span data-stu-id="8639f-753">Error code</span></span>|<span data-ttu-id="8639f-754">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-754">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="8639f-755">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="8639f-755">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="8639f-756">許可されていない拡張子付きの添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="8639f-756">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="8639f-757">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="8639f-757">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8639f-758">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-758">Requirements</span></span>

|<span data-ttu-id="8639f-759">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-759">Requirement</span></span>|<span data-ttu-id="8639f-760">値</span><span class="sxs-lookup"><span data-stu-id="8639f-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-761">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-761">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-762">プレビュー</span><span class="sxs-lookup"><span data-stu-id="8639f-762">Preview</span></span>|
|[<span data-ttu-id="8639f-763">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-763">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-764">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8639f-764">ReadWriteItem</span></span>|
|[<span data-ttu-id="8639f-765">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-765">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-766">新規作成</span><span class="sxs-lookup"><span data-stu-id="8639f-766">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8639f-767">例</span><span class="sxs-lookup"><span data-stu-id="8639f-767">Examples</span></span>

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="8639f-768">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8639f-768">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="8639f-769">サポートされているイベントのイベント ハンドラを追加します。</span><span class="sxs-lookup"><span data-stu-id="8639f-769">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="8639f-770">現在、サポートされているイベントの種類は、`Office.EventType.AppointmentTimeChanged`と`Office.EventType.RecipientsChanged`です。 `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="8639f-770">Currently, the supported event types are `Office.EventType.AppointmentTimeChanged` and `Office.EventType.RecipientsChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8639f-771">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="8639f-771">Parameters:</span></span>

| <span data-ttu-id="8639f-772">名前</span><span class="sxs-lookup"><span data-stu-id="8639f-772">Name</span></span> | <span data-ttu-id="8639f-773">種類</span><span class="sxs-lookup"><span data-stu-id="8639f-773">Type</span></span> | <span data-ttu-id="8639f-774">属性</span><span class="sxs-lookup"><span data-stu-id="8639f-774">Attributes</span></span> | <span data-ttu-id="8639f-775">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-775">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="8639f-776">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="8639f-776">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="8639f-777">ハンドラを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="8639f-777">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="8639f-778">関数</span><span class="sxs-lookup"><span data-stu-id="8639f-778">Function</span></span> || <span data-ttu-id="8639f-p138">イベントを処理する関数。この関数は、オブジェクト リテラルである単一パラメータを受け入れる必要があります。パラメータの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメータと一致します。</span><span class="sxs-lookup"><span data-stu-id="8639f-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="8639f-782">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-782">Object</span></span> | <span data-ttu-id="8639f-783">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-783">&lt;optional&gt;</span></span> | <span data-ttu-id="8639f-784">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="8639f-784">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="8639f-785">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-785">Object</span></span> | <span data-ttu-id="8639f-786">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-786">&lt;optional&gt;</span></span> | <span data-ttu-id="8639f-787">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8639f-787">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="8639f-788">関数</span><span class="sxs-lookup"><span data-stu-id="8639f-788">function</span></span>| <span data-ttu-id="8639f-789">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-789">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-790">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="8639f-790">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8639f-791">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-791">Requirements</span></span>

|<span data-ttu-id="8639f-792">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-792">Requirement</span></span>| <span data-ttu-id="8639f-793">値</span><span class="sxs-lookup"><span data-stu-id="8639f-793">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-794">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-794">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8639f-795">1.7</span><span class="sxs-lookup"><span data-stu-id="8639f-795">-17</span></span> |
|[<span data-ttu-id="8639f-796">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-796">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8639f-797">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-797">ReadItem</span></span> |
|[<span data-ttu-id="8639f-798">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-798">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8639f-799">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8639f-799">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="8639f-800">addItemAttachmentAsync (itemd、attachmentName、[オプション]、[コールバック])</span><span class="sxs-lookup"><span data-stu-id="8639f-800">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8639f-801">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="8639f-801">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="8639f-p139">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つ項目を新規作成フォーム内の項目に添付します。コールバック メソッドを指定する場合、`asyncResult` というパラメータがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、または項目を添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメータを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="8639f-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="8639f-805">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="8639f-805">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="8639f-806">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドは項目を、編集中以外の項目に添付できますが、これはサポートされておらず、推奨されていません。</span><span class="sxs-lookup"><span data-stu-id="8639f-806">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8639f-807">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="8639f-807">Parameters:</span></span>

|<span data-ttu-id="8639f-808">名前</span><span class="sxs-lookup"><span data-stu-id="8639f-808">Name</span></span>|<span data-ttu-id="8639f-809">種類</span><span class="sxs-lookup"><span data-stu-id="8639f-809">Type</span></span>|<span data-ttu-id="8639f-810">属性</span><span class="sxs-lookup"><span data-stu-id="8639f-810">Attributes</span></span>|<span data-ttu-id="8639f-811">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-811">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="8639f-812">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-812">String</span></span>||<span data-ttu-id="8639f-p140">添付するアイテムの Exchange 識別子です。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="8639f-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="8639f-815">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-815">String</span></span>||<span data-ttu-id="8639f-p141">添付するアイテムの件名です。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="8639f-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="8639f-818">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-818">Object</span></span>|<span data-ttu-id="8639f-819">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-819">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-820">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="8639f-820">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8639f-821">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-821">Object</span></span>|<span data-ttu-id="8639f-822">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-822">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-823">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8639f-823">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8639f-824">関数</span><span class="sxs-lookup"><span data-stu-id="8639f-824">function</span></span>|<span data-ttu-id="8639f-825">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-825">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-826">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="8639f-826">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8639f-827">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-827">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8639f-828">添付ファイルの追加に失敗した場合、 エラーの説明を提供する`asyncResult` オブジェクトが `Error` オブジェクトに含まれます。</span><span class="sxs-lookup"><span data-stu-id="8639f-828">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8639f-829">エラー</span><span class="sxs-lookup"><span data-stu-id="8639f-829">Errors</span></span>

|<span data-ttu-id="8639f-830">エラー コード</span><span class="sxs-lookup"><span data-stu-id="8639f-830">Error code</span></span>|<span data-ttu-id="8639f-831">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-831">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="8639f-832">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="8639f-832">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8639f-833">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-833">Requirements</span></span>

|<span data-ttu-id="8639f-834">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-834">Requirement</span></span>|<span data-ttu-id="8639f-835">値</span><span class="sxs-lookup"><span data-stu-id="8639f-835">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-836">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-836">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-837">1.1</span><span class="sxs-lookup"><span data-stu-id="8639f-837">1.1</span></span>|
|[<span data-ttu-id="8639f-838">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-838">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-839">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8639f-839">ReadWriteItem</span></span>|
|[<span data-ttu-id="8639f-840">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-840">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-841">新規作成</span><span class="sxs-lookup"><span data-stu-id="8639f-841">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-842">例</span><span class="sxs-lookup"><span data-stu-id="8639f-842">Example</span></span>

<span data-ttu-id="8639f-843">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-843">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="8639f-844">閉じる()</span><span class="sxs-lookup"><span data-stu-id="8639f-844">close()</span></span>

<span data-ttu-id="8639f-845">新規作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="8639f-845">Closes the current item that is being composed.</span></span>

<span data-ttu-id="8639f-p142">`close`メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="8639f-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="8639f-848">Outlook on the webでは、項目が予定で、`saveAsync`を用いて事前に保存されている場合、項目が最後に保存されてから何も変更されていない場合でも、ユーザーに対して保存、破棄またはキャンセルするよう求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="8639f-848">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="8639f-849">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close`メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="8639f-849">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-850">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-850">Requirements</span></span>

|<span data-ttu-id="8639f-851">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-851">Requirement</span></span>|<span data-ttu-id="8639f-852">値</span><span class="sxs-lookup"><span data-stu-id="8639f-852">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-853">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-853">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-854">1.3</span><span class="sxs-lookup"><span data-stu-id="8639f-854">1.3</span></span>|
|[<span data-ttu-id="8639f-855">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-855">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-856">制限あり</span><span class="sxs-lookup"><span data-stu-id="8639f-856">Restricted</span></span>|
|[<span data-ttu-id="8639f-857">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-857">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-858">新規作成</span><span class="sxs-lookup"><span data-stu-id="8639f-858">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="8639f-859">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="8639f-859">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="8639f-860">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-860">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8639f-861">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8639f-861">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8639f-862">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-862">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8639f-863">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="8639f-863">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="8639f-p143">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="8639f-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8639f-867">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="8639f-867">Parameters:</span></span>

|<span data-ttu-id="8639f-868">名前</span><span class="sxs-lookup"><span data-stu-id="8639f-868">Name</span></span>|<span data-ttu-id="8639f-869">種類</span><span class="sxs-lookup"><span data-stu-id="8639f-869">Type</span></span>|<span data-ttu-id="8639f-870">属性</span><span class="sxs-lookup"><span data-stu-id="8639f-870">Attributes</span></span>|<span data-ttu-id="8639f-871">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-871">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="8639f-872">文字列 |オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-872">String &#124; Object</span></span>||<span data-ttu-id="8639f-p144">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="8639f-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8639f-875">**または**</span><span class="sxs-lookup"><span data-stu-id="8639f-875">**OR**</span></span><br/><span data-ttu-id="8639f-p145">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクトです。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8639f-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="8639f-878">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-878">String</span></span>|<span data-ttu-id="8639f-879">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-879">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="8639f-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="8639f-882">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-882">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="8639f-883">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-883">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-884">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="8639f-884">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="8639f-885">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-885">String</span></span>||<span data-ttu-id="8639f-p147">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` でなければいけません。</span><span class="sxs-lookup"><span data-stu-id="8639f-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="8639f-888">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-888">String</span></span>||<span data-ttu-id="8639f-889">添付ファイル名を含む文字列です。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="8639f-889">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="8639f-890">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-890">String</span></span>||<span data-ttu-id="8639f-p148">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URIです。</span><span class="sxs-lookup"><span data-stu-id="8639f-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="8639f-893">ブール値</span><span class="sxs-lookup"><span data-stu-id="8639f-893">Boolean</span></span>||<span data-ttu-id="8639f-p149">`type`が`file`に設定されている場合にのみ使用されます。`true`の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="8639f-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="8639f-896">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-896">String</span></span>||<span data-ttu-id="8639f-p150">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムIDです。最大長が 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="8639f-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="8639f-900">関数</span><span class="sxs-lookup"><span data-stu-id="8639f-900">function</span></span>|<span data-ttu-id="8639f-901">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-901">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-902">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである シングル パラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-902">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8639f-903">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-903">Requirements</span></span>

|<span data-ttu-id="8639f-904">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-904">Requirement</span></span>|<span data-ttu-id="8639f-905">値</span><span class="sxs-lookup"><span data-stu-id="8639f-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-906">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-907">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-907">1.0</span></span>|
|[<span data-ttu-id="8639f-908">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-909">ReadItem</span></span>|
|[<span data-ttu-id="8639f-910">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-911">読み取り</span><span class="sxs-lookup"><span data-stu-id="8639f-911">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8639f-912">例</span><span class="sxs-lookup"><span data-stu-id="8639f-912">Examples</span></span>

<span data-ttu-id="8639f-913">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="8639f-913">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="8639f-914">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="8639f-914">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="8639f-915">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="8639f-915">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8639f-916">本文と添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="8639f-916">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="8639f-917">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="8639f-917">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="8639f-918">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="8639f-918">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="8639f-919">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="8639f-919">displayReplyForm(formData)</span></span>

<span data-ttu-id="8639f-920">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む返信フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-920">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8639f-921">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8639f-921">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8639f-922">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-922">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8639f-923">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="8639f-923">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="8639f-p151">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="8639f-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8639f-927">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="8639f-927">Parameters:</span></span>

|<span data-ttu-id="8639f-928">名前</span><span class="sxs-lookup"><span data-stu-id="8639f-928">Name</span></span>|<span data-ttu-id="8639f-929">種類</span><span class="sxs-lookup"><span data-stu-id="8639f-929">Type</span></span>|<span data-ttu-id="8639f-930">属性</span><span class="sxs-lookup"><span data-stu-id="8639f-930">Attributes</span></span>|<span data-ttu-id="8639f-931">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-931">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="8639f-932">文字列 |オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-932">String &#124; Object</span></span>||<span data-ttu-id="8639f-p152">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="8639f-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8639f-935">**または**</span><span class="sxs-lookup"><span data-stu-id="8639f-935">**OR**</span></span><br/><span data-ttu-id="8639f-p153">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクトです。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8639f-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="8639f-938">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-938">String</span></span>|<span data-ttu-id="8639f-939">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-939">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-p154">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="8639f-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="8639f-942">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-942">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="8639f-943">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-943">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-944">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="8639f-944">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="8639f-945">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-945">String</span></span>||<span data-ttu-id="8639f-p155">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` でなければいけません。</span><span class="sxs-lookup"><span data-stu-id="8639f-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="8639f-948">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-948">String</span></span>||<span data-ttu-id="8639f-949">添付ファイル名を含む文字列です。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="8639f-949">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="8639f-950">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-950">String</span></span>||<span data-ttu-id="8639f-p156">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URIです。</span><span class="sxs-lookup"><span data-stu-id="8639f-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="8639f-953">ブール値</span><span class="sxs-lookup"><span data-stu-id="8639f-953">Boolean</span></span>||<span data-ttu-id="8639f-p157">`type`が`file`に設定されている場合にのみ使用されます。`true`の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="8639f-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="8639f-956">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-956">String</span></span>||<span data-ttu-id="8639f-p158">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムIDです。最大長が 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="8639f-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="8639f-960">関数</span><span class="sxs-lookup"><span data-stu-id="8639f-960">function</span></span>|<span data-ttu-id="8639f-961">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-961">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-962">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである シングル パラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-962">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8639f-963">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-963">Requirements</span></span>

|<span data-ttu-id="8639f-964">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-964">Requirement</span></span>|<span data-ttu-id="8639f-965">値</span><span class="sxs-lookup"><span data-stu-id="8639f-965">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-966">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-966">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-967">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-967">1.0</span></span>|
|[<span data-ttu-id="8639f-968">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-968">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-969">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-969">ReadItem</span></span>|
|[<span data-ttu-id="8639f-970">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-970">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-971">読み取り</span><span class="sxs-lookup"><span data-stu-id="8639f-971">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8639f-972">例</span><span class="sxs-lookup"><span data-stu-id="8639f-972">Examples</span></span>

<span data-ttu-id="8639f-973">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="8639f-973">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="8639f-974">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="8639f-974">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="8639f-975">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="8639f-975">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8639f-976">本文と添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="8639f-976">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="8639f-977">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="8639f-977">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="8639f-978">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="8639f-978">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="8639f-979">getEntities() → {[エンティティ](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="8639f-979">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="8639f-980">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="8639f-980">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="8639f-981">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8639f-981">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-982">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-982">Requirements</span></span>

|<span data-ttu-id="8639f-983">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-983">Requirement</span></span>|<span data-ttu-id="8639f-984">値</span><span class="sxs-lookup"><span data-stu-id="8639f-984">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-985">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-985">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-986">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-986">1.0</span></span>|
|[<span data-ttu-id="8639f-987">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-987">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-988">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-988">ReadItem</span></span>|
|[<span data-ttu-id="8639f-989">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-989">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-990">読み取り</span><span class="sxs-lookup"><span data-stu-id="8639f-990">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8639f-991">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="8639f-991">Returns:</span></span>

<span data-ttu-id="8639f-992">種類: [エンティティ](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="8639f-992">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="8639f-993">例</span><span class="sxs-lookup"><span data-stu-id="8639f-993">Example</span></span>

<span data-ttu-id="8639f-994">次の例では、現在の項目の本文内にある連絡先のエンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="8639f-994">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="8639f-995">getEntitiesByType(entityType)] → [(空白可能) {<(String|[連絡先](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[電話番号](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)) >}</span><span class="sxs-lookup"><span data-stu-id="8639f-995">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="8639f-996">選択した項目で見つかった指定のエンティティ型のエンティティすべてを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="8639f-996">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="8639f-997">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8639f-997">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8639f-998">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="8639f-998">Parameters:</span></span>

|<span data-ttu-id="8639f-999">名前</span><span class="sxs-lookup"><span data-stu-id="8639f-999">Name</span></span>|<span data-ttu-id="8639f-1000">種類</span><span class="sxs-lookup"><span data-stu-id="8639f-1000">Type</span></span>|<span data-ttu-id="8639f-1001">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-1001">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="8639f-1002">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="8639f-1002">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="8639f-1003">EntityType 列挙値の 1 つです。</span><span class="sxs-lookup"><span data-stu-id="8639f-1003">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8639f-1004">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1004">Requirements</span></span>

|<span data-ttu-id="8639f-1005">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1005">Requirement</span></span>|<span data-ttu-id="8639f-1006">値</span><span class="sxs-lookup"><span data-stu-id="8639f-1006">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-1007">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-1007">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-1008">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-1008">1.0</span></span>|
|[<span data-ttu-id="8639f-1009">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-1009">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-1010">制限あり</span><span class="sxs-lookup"><span data-stu-id="8639f-1010">Restricted</span></span>|
|[<span data-ttu-id="8639f-1011">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-1011">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-1012">読み取り</span><span class="sxs-lookup"><span data-stu-id="8639f-1012">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8639f-1013">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="8639f-1013">Returns:</span></span>

<span data-ttu-id="8639f-1014">`entityType` に渡された値が有効な `EntityType` 列挙型のメンバーでない場合、メソッドは 空白を返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-1014">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="8639f-1015">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-1015">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="8639f-1016">そうでない場合、返される配列内のオブジェクトの種類は、 `entityType` パラメータ内で要求されたエンティティの種類によって異なります。</span><span class="sxs-lookup"><span data-stu-id="8639f-1016">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="8639f-1017">このメソッドを使用する最小限のアクセス許可レベルは **制限あり** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="8639f-1017">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="8639f-1018">の値 `entityType`</span><span class="sxs-lookup"><span data-stu-id="8639f-1018">Value of `entityType`</span></span>|<span data-ttu-id="8639f-1019">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="8639f-1019">Type of objects in returned array</span></span>|<span data-ttu-id="8639f-1020">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="8639f-1020">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="8639f-1021">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-1021">String</span></span>|<span data-ttu-id="8639f-1022">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="8639f-1022">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="8639f-1023">連絡先</span><span class="sxs-lookup"><span data-stu-id="8639f-1023">Contact</span></span>|<span data-ttu-id="8639f-1024">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8639f-1024">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="8639f-1025">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-1025">String</span></span>|<span data-ttu-id="8639f-1026">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8639f-1026">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="8639f-1027">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="8639f-1027">MeetingSuggestion</span></span>|<span data-ttu-id="8639f-1028">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8639f-1028">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="8639f-1029">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="8639f-1029">PhoneNumber</span></span>|<span data-ttu-id="8639f-1030">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="8639f-1030">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="8639f-1031">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="8639f-1031">TaskSuggestion</span></span>|<span data-ttu-id="8639f-1032">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8639f-1032">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="8639f-1033">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-1033">String</span></span>|<span data-ttu-id="8639f-1034">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="8639f-1034">**Restricted**</span></span>|

<span data-ttu-id="8639f-1035">型:Array.<(文字列|[連絡先](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="8639f-1035">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="8639f-1036">例</span><span class="sxs-lookup"><span data-stu-id="8639f-1036">Example</span></span>

<span data-ttu-id="8639f-1037">次の例は、現在のアイテムの本文にある郵便アドレスを表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="8639f-1037">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="8639f-1038">getFilteredEntitiesByName(name)] → [(Null 許容) {<(文字列| [連絡先](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[電話番号 ](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="8639f-1038">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="8639f-1039">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-1039">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8639f-1040">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8639f-1040">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8639f-1041">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-1041">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8639f-1042">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="8639f-1042">Parameters:</span></span>

|<span data-ttu-id="8639f-1043">名前</span><span class="sxs-lookup"><span data-stu-id="8639f-1043">Name</span></span>|<span data-ttu-id="8639f-1044">種類</span><span class="sxs-lookup"><span data-stu-id="8639f-1044">Type</span></span>|<span data-ttu-id="8639f-1045">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-1045">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="8639f-1046">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-1046">String</span></span>|<span data-ttu-id="8639f-1047">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前です。</span><span class="sxs-lookup"><span data-stu-id="8639f-1047">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8639f-1048">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1048">Requirements</span></span>

|<span data-ttu-id="8639f-1049">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1049">Requirement</span></span>|<span data-ttu-id="8639f-1050">値</span><span class="sxs-lookup"><span data-stu-id="8639f-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-1051">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-1051">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-1052">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-1052">1.0</span></span>|
|[<span data-ttu-id="8639f-1053">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-1053">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-1054">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-1054">ReadItem</span></span>|
|[<span data-ttu-id="8639f-1055">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-1055">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-1056">読み取り</span><span class="sxs-lookup"><span data-stu-id="8639f-1056">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8639f-1057">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="8639f-1057">Returns:</span></span>

<span data-ttu-id="8639f-p160">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="8639f-1060">型:Array.<(文字列|[連絡先](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="8639f-1060">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="8639f-1061">getInitializationContextAsync ([オプション]、[コールバック])</span><span class="sxs-lookup"><span data-stu-id="8639f-1061">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="8639f-1062">アドインが[操作可能メッセージによってアクティブ化](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを取得します。</span><span class="sxs-lookup"><span data-stu-id="8639f-1062">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="8639f-1063">注:このメソッドは、Outlook 2016 for Windows (16.0.8413.1000 以降のクイック実行バージョン) および Outlook on the web for Office 365 でのみサポートされます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1063">Note: This method is only supported by Outlook 2016 for Windows (Click-to-Run versions greater than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8639f-1064">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="8639f-1064">Parameters:</span></span>
|<span data-ttu-id="8639f-1065">名前</span><span class="sxs-lookup"><span data-stu-id="8639f-1065">Name</span></span>|<span data-ttu-id="8639f-1066">種類</span><span class="sxs-lookup"><span data-stu-id="8639f-1066">Type</span></span>|<span data-ttu-id="8639f-1067">属性</span><span class="sxs-lookup"><span data-stu-id="8639f-1067">Attributes</span></span>|<span data-ttu-id="8639f-1068">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-1068">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="8639f-1069">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-1069">Object</span></span>|<span data-ttu-id="8639f-1070">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-1070">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-1071">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="8639f-1071">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8639f-1072">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-1072">Object</span></span>|<span data-ttu-id="8639f-1073">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-1073">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-1074">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1074">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8639f-1075">関数</span><span class="sxs-lookup"><span data-stu-id="8639f-1075">function</span></span>|<span data-ttu-id="8639f-1076">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-1076">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-1077">メソッドが完了すると、`callback` パラメータで渡された関数が、単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="8639f-1077">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8639f-1078">成功すると、初期化データが文字列として `asyncResult.value` プロパティで指定されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1078">On success, the intialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="8639f-1079">初期化コンテキストがない場合、`asyncResult` オブジェクトには、`code` プロパティが `9020`、`name` プロパティが `GenericResponseError` に設定された `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1079">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8639f-1080">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1080">Requirements</span></span>

|<span data-ttu-id="8639f-1081">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1081">Requirement</span></span>|<span data-ttu-id="8639f-1082">値</span><span class="sxs-lookup"><span data-stu-id="8639f-1082">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-1083">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-1083">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-1084">プレビュー</span><span class="sxs-lookup"><span data-stu-id="8639f-1084">Preview</span></span>|
|[<span data-ttu-id="8639f-1085">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-1085">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-1086">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-1086">ReadItem</span></span>|
|[<span data-ttu-id="8639f-1087">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-1087">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-1088">読み取り</span><span class="sxs-lookup"><span data-stu-id="8639f-1088">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-1089">例</span><span class="sxs-lookup"><span data-stu-id="8639f-1089">Example</span></span>

```
// Get the initialization context (if present)
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object
        var context = JSON.parse(asyncResult.value);
        // Do something with context
      } else {
        // Empty context, treat as no context
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is
        // no context
        // Treat as no context
      } else {
        // Handle the error
      }
    }
  }
);
```

#### <a name="getregexmatches--object"></a><span data-ttu-id="8639f-1090">getRegExMatches() → {オブジェクト}</span><span class="sxs-lookup"><span data-stu-id="8639f-1090">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="8639f-1091">選択した項目内で、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-1091">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8639f-1092">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8639f-1092">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8639f-p161">`getRegExMatches` メソッドは、マニフェスト XML ファイル内にある各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定された項目のプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="8639f-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="8639f-1096">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="8639f-1096">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="8639f-1097">`getRegExMatches` から返されるオブジェクトには、 `fruits` および `veggies` という 2 つのプロパティがありえます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1097">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="8639f-p162">アイテムの 本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="8639f-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-1101">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1101">Requirements</span></span>

|<span data-ttu-id="8639f-1102">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1102">Requirement</span></span>|<span data-ttu-id="8639f-1103">値</span><span class="sxs-lookup"><span data-stu-id="8639f-1103">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-1104">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-1104">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-1105">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-1105">1.0</span></span>|
|[<span data-ttu-id="8639f-1106">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-1106">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-1107">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-1107">ReadItem</span></span>|
|[<span data-ttu-id="8639f-1108">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-1108">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-1109">読み取り</span><span class="sxs-lookup"><span data-stu-id="8639f-1109">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8639f-1110">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="8639f-1110">Returns:</span></span>

<span data-ttu-id="8639f-p163">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクトです。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="8639f-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="8639f-1113">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="8639f-1113">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8639f-1114">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-1114">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8639f-1115">例</span><span class="sxs-lookup"><span data-stu-id="8639f-1115">Example</span></span>

<span data-ttu-id="8639f-1116">次の例は、マニフェストで指定された正規表現ルールの要素`fruits`および`veggies`に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8639f-1116">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="8639f-1117">getRegExMatchesByName(name)] → [(Null 許容) {配列.< 文字列 >}</span><span class="sxs-lookup"><span data-stu-id="8639f-1117">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="8639f-1118">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-1118">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8639f-1119">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8639f-1119">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8639f-1120">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="8639f-1120">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="8639f-p164">項目の 本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、項目の本文全体を返さないようにします。`.*` などの正規表現を使用して項目の本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="8639f-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8639f-1123">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="8639f-1123">Parameters:</span></span>

|<span data-ttu-id="8639f-1124">名前</span><span class="sxs-lookup"><span data-stu-id="8639f-1124">Name</span></span>|<span data-ttu-id="8639f-1125">種類</span><span class="sxs-lookup"><span data-stu-id="8639f-1125">Type</span></span>|<span data-ttu-id="8639f-1126">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-1126">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="8639f-1127">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-1127">String</span></span>|<span data-ttu-id="8639f-1128">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前です。</span><span class="sxs-lookup"><span data-stu-id="8639f-1128">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8639f-1129">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1129">Requirements</span></span>

|<span data-ttu-id="8639f-1130">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1130">Requirement</span></span>|<span data-ttu-id="8639f-1131">値</span><span class="sxs-lookup"><span data-stu-id="8639f-1131">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-1132">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-1132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-1133">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-1133">1.0</span></span>|
|[<span data-ttu-id="8639f-1134">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-1134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-1135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-1135">ReadItem</span></span>|
|[<span data-ttu-id="8639f-1136">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-1136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-1137">読み取り</span><span class="sxs-lookup"><span data-stu-id="8639f-1137">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8639f-1138">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="8639f-1138">Returns:</span></span>

<span data-ttu-id="8639f-1139">マニフェスト XML ファイルで定義された正規表現に一致する文字列を含む配列です。</span><span class="sxs-lookup"><span data-stu-id="8639f-1139">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="8639f-1140">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="8639f-1140">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8639f-1141">配列. < 文字列 ></span><span class="sxs-lookup"><span data-stu-id="8639f-1141">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8639f-1142">例</span><span class="sxs-lookup"><span data-stu-id="8639f-1142">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="8639f-1143">getSelectedDataAsync(coercionType, [オプション], コールバック) → {文字列}</span><span class="sxs-lookup"><span data-stu-id="8639f-1143">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="8639f-1144">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-1144">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="8639f-p165">選択したデータがなく、カーソルが本文または件名にある場合、メソッドは選択したデータに対して Null を返します。本文または件名以外のフィールドが選択されている場合、メソッドは `InvalidSelection` エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8639f-1147">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="8639f-1147">Parameters:</span></span>

|<span data-ttu-id="8639f-1148">名前</span><span class="sxs-lookup"><span data-stu-id="8639f-1148">Name</span></span>|<span data-ttu-id="8639f-1149">種類</span><span class="sxs-lookup"><span data-stu-id="8639f-1149">Type</span></span>|<span data-ttu-id="8639f-1150">属性</span><span class="sxs-lookup"><span data-stu-id="8639f-1150">Attributes</span></span>|<span data-ttu-id="8639f-1151">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-1151">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="8639f-1152">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8639f-1152">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="8639f-p166">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="8639f-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="8639f-1156">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-1156">Object</span></span>|<span data-ttu-id="8639f-1157">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-1158">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="8639f-1158">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8639f-1159">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-1159">Object</span></span>|<span data-ttu-id="8639f-1160">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-1161">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1161">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8639f-1162">関数</span><span class="sxs-lookup"><span data-stu-id="8639f-1162">function</span></span>||<span data-ttu-id="8639f-1163">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="8639f-1163">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8639f-1164">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data`を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="8639f-1164">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="8639f-1165">選択元のソース プロパティにアクセスするには、 `asyncResult.value.sourceProperty` を呼び出します。これは `body`   または `subject`   になります。</span><span class="sxs-lookup"><span data-stu-id="8639f-1165">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8639f-1166">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1166">Requirements</span></span>

|<span data-ttu-id="8639f-1167">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1167">Requirement</span></span>|<span data-ttu-id="8639f-1168">値</span><span class="sxs-lookup"><span data-stu-id="8639f-1168">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-1169">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-1169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-1170">1.2</span><span class="sxs-lookup"><span data-stu-id="8639f-1170">1.2</span></span>|
|[<span data-ttu-id="8639f-1171">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-1171">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-1172">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8639f-1172">ReadWriteItem</span></span>|
|[<span data-ttu-id="8639f-1173">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-1173">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-1174">新規作成</span><span class="sxs-lookup"><span data-stu-id="8639f-1174">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="8639f-1175">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="8639f-1175">Returns:</span></span>

<span data-ttu-id="8639f-1176">`coercionType`で決定された書式設定の文字列として選択されたデータです。</span><span class="sxs-lookup"><span data-stu-id="8639f-1176">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="8639f-1177">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="8639f-1177">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8639f-1178">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-1178">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8639f-1179">例</span><span class="sxs-lookup"><span data-stu-id="8639f-1179">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="8639f-1180">getSelectedEntities() → {[エンティティ](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="8639f-1180">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="8639f-p168">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-p168">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="8639f-1183">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8639f-1183">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-1184">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1184">Requirements</span></span>

|<span data-ttu-id="8639f-1185">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1185">Requirement</span></span>|<span data-ttu-id="8639f-1186">値</span><span class="sxs-lookup"><span data-stu-id="8639f-1186">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-1187">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-1187">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-1188">1.6</span><span class="sxs-lookup"><span data-stu-id="8639f-1188">-16</span></span>|
|[<span data-ttu-id="8639f-1189">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-1189">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-1190">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-1190">ReadItem</span></span>|
|[<span data-ttu-id="8639f-1191">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-1191">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-1192">読み取り</span><span class="sxs-lookup"><span data-stu-id="8639f-1192">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8639f-1193">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="8639f-1193">Returns:</span></span>

<span data-ttu-id="8639f-1194">種類: [エンティティ](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="8639f-1194">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="8639f-1195">例</span><span class="sxs-lookup"><span data-stu-id="8639f-1195">Example</span></span>

<span data-ttu-id="8639f-1196">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="8639f-1196">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="8639f-1197">getSelectedRegExMatches() → {オブジェクト}</span><span class="sxs-lookup"><span data-stu-id="8639f-1197">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="8639f-p169">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="8639f-1200">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8639f-1200">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8639f-p170">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内にある各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定された項目のプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="8639f-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="8639f-1204">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="8639f-1204">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="8639f-1205">`getRegExMatches` から返されるオブジェクトには、 `fruits` および `veggies` という 2 つのプロパティがありえます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1205">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="8639f-p171">アイテムの 本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="8639f-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8639f-1209">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1209">Requirements</span></span>

|<span data-ttu-id="8639f-1210">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1210">Requirement</span></span>|<span data-ttu-id="8639f-1211">値</span><span class="sxs-lookup"><span data-stu-id="8639f-1211">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-1212">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-1212">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-1213">1.6</span><span class="sxs-lookup"><span data-stu-id="8639f-1213">-16</span></span>|
|[<span data-ttu-id="8639f-1214">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-1214">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-1215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-1215">ReadItem</span></span>|
|[<span data-ttu-id="8639f-1216">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-1216">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-1217">読み取り</span><span class="sxs-lookup"><span data-stu-id="8639f-1217">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8639f-1218">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="8639f-1218">Returns:</span></span>

<span data-ttu-id="8639f-p172">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクトです。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="8639f-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="8639f-1221">例</span><span class="sxs-lookup"><span data-stu-id="8639f-1221">Example</span></span>

<span data-ttu-id="8639f-1222">次の例は、マニフェストで指定された正規表現ルールの要素`fruits`および`veggies`に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8639f-1222">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="8639f-1223">getSharedPropertiesAsync ([オプション]、コールバック)</span><span class="sxs-lookup"><span data-stu-id="8639f-1223">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="8639f-1224">共有フォルダー、予定表、またはメールボックス内の選択されている予定またはメッセージのプロパティを取得します。</span><span class="sxs-lookup"><span data-stu-id="8639f-1224">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8639f-1225">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="8639f-1225">Parameters:</span></span>

|<span data-ttu-id="8639f-1226">名前</span><span class="sxs-lookup"><span data-stu-id="8639f-1226">Name</span></span>|<span data-ttu-id="8639f-1227">種類</span><span class="sxs-lookup"><span data-stu-id="8639f-1227">Type</span></span>|<span data-ttu-id="8639f-1228">属性</span><span class="sxs-lookup"><span data-stu-id="8639f-1228">Attributes</span></span>|<span data-ttu-id="8639f-1229">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-1229">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="8639f-1230">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-1230">Object</span></span>|<span data-ttu-id="8639f-1231">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-1231">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-1232">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="8639f-1232">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8639f-1233">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-1233">Object</span></span>|<span data-ttu-id="8639f-1234">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-1234">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-1235">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1235">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8639f-1236">関数</span><span class="sxs-lookup"><span data-stu-id="8639f-1236">function</span></span>||<span data-ttu-id="8639f-1237">メソッドが完了すると、`callback` パラメータで渡された関数が、単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="8639f-1237">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8639f-1238">共有のプロパティは `asyncResult.value` プロパティの [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1238">The custom properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="8639f-1239">このオブジェクトは、アイテムの共有のプロパティの取得に使用できます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1239">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8639f-1240">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1240">Requirements</span></span>

|<span data-ttu-id="8639f-1241">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1241">Requirement</span></span>|<span data-ttu-id="8639f-1242">値</span><span class="sxs-lookup"><span data-stu-id="8639f-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-1243">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-1243">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-1244">プレビュー</span><span class="sxs-lookup"><span data-stu-id="8639f-1244">Preview</span></span>|
|[<span data-ttu-id="8639f-1245">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-1245">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-1246">ReadItem</span></span>|
|[<span data-ttu-id="8639f-1247">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-1247">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-1248">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8639f-1248">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-1249">例</span><span class="sxs-lookup"><span data-stu-id="8639f-1249">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="8639f-1250">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8639f-1250">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="8639f-1251">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1251">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="8639f-p174">カスタム プロパティは、アプリケーションごと、アイテムごとにキーと値のペアとして保管されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在の項目および現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、項目上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="8639f-p174">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8639f-1255">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="8639f-1255">Parameters:</span></span>

|<span data-ttu-id="8639f-1256">名前</span><span class="sxs-lookup"><span data-stu-id="8639f-1256">Name</span></span>|<span data-ttu-id="8639f-1257">種類</span><span class="sxs-lookup"><span data-stu-id="8639f-1257">Type</span></span>|<span data-ttu-id="8639f-1258">属性</span><span class="sxs-lookup"><span data-stu-id="8639f-1258">Attributes</span></span>|<span data-ttu-id="8639f-1259">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-1259">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="8639f-1260">関数</span><span class="sxs-lookup"><span data-stu-id="8639f-1260">function</span></span>||<span data-ttu-id="8639f-1261">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="8639f-1261">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8639f-1262">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1262">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="8639f-1263">項目からカスタム プロパティを取得、設定、削除して、サーバーにカスタム プロパティのセット バックに対する変更を保存するのに、このオブジェクトが使用できます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1263">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="8639f-1264">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-1264">Object</span></span>|<span data-ttu-id="8639f-1265">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-1265">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-1266">開発者は、コールバック 関数でアクセスしたいオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1266">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="8639f-1267">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1267">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8639f-1268">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1268">Requirements</span></span>

|<span data-ttu-id="8639f-1269">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1269">Requirement</span></span>|<span data-ttu-id="8639f-1270">値</span><span class="sxs-lookup"><span data-stu-id="8639f-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-1271">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-1272">1.0</span><span class="sxs-lookup"><span data-stu-id="8639f-1272">1.0</span></span>|
|[<span data-ttu-id="8639f-1273">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-1273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-1274">ReadItem</span></span>|
|[<span data-ttu-id="8639f-1275">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-1275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-1276">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8639f-1276">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-1277">例</span><span class="sxs-lookup"><span data-stu-id="8639f-1277">Example</span></span>

<span data-ttu-id="8639f-p177">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="8639f-p177">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="8639f-1281">removeAttachmentAsync(attachmentId、[オプション]、 [コールバック])</span><span class="sxs-lookup"><span data-stu-id="8639f-1281">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="8639f-1282">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="8639f-1282">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="8639f-p178">`removeAttachmentAsync`メソッドは、指定した識別子の添付ファイルを項目 から削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web アプリ とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別のウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="8639f-p178">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8639f-1287">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="8639f-1287">Parameters:</span></span>

|<span data-ttu-id="8639f-1288">名前</span><span class="sxs-lookup"><span data-stu-id="8639f-1288">Name</span></span>|<span data-ttu-id="8639f-1289">種類</span><span class="sxs-lookup"><span data-stu-id="8639f-1289">Type</span></span>|<span data-ttu-id="8639f-1290">属性</span><span class="sxs-lookup"><span data-stu-id="8639f-1290">Attributes</span></span>|<span data-ttu-id="8639f-1291">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-1291">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="8639f-1292">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-1292">String</span></span>||<span data-ttu-id="8639f-p179">削除する添付ファイルの識別子です。文字列の最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="8639f-p179">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="8639f-1295">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-1295">Object</span></span>|<span data-ttu-id="8639f-1296">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-1296">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-1297">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="8639f-1297">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8639f-1298">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-1298">Object</span></span>|<span data-ttu-id="8639f-1299">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-1299">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-1300">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1300">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8639f-1301">関数</span><span class="sxs-lookup"><span data-stu-id="8639f-1301">function</span></span>|<span data-ttu-id="8639f-1302">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-1303">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="8639f-1303">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8639f-1304">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1304">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8639f-1305">エラー</span><span class="sxs-lookup"><span data-stu-id="8639f-1305">Errors</span></span>

|<span data-ttu-id="8639f-1306">エラー コード</span><span class="sxs-lookup"><span data-stu-id="8639f-1306">Error code</span></span>|<span data-ttu-id="8639f-1307">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-1307">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="8639f-1308">添付ファイルの識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="8639f-1308">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8639f-1309">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1309">Requirements</span></span>

|<span data-ttu-id="8639f-1310">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1310">Requirement</span></span>|<span data-ttu-id="8639f-1311">値</span><span class="sxs-lookup"><span data-stu-id="8639f-1311">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-1312">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-1312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-1313">1.1</span><span class="sxs-lookup"><span data-stu-id="8639f-1313">1.1</span></span>|
|[<span data-ttu-id="8639f-1314">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-1314">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-1315">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8639f-1315">ReadWriteItem</span></span>|
|[<span data-ttu-id="8639f-1316">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-1316">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-1317">新規作成</span><span class="sxs-lookup"><span data-stu-id="8639f-1317">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-1318">例</span><span class="sxs-lookup"><span data-stu-id="8639f-1318">Example</span></span>

<span data-ttu-id="8639f-1319">次のコードは、「0」の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="8639f-1319">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="8639f-1320">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8639f-1320">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="8639f-1321">サポートされているイベントのイベント ハンドラを追加します。</span><span class="sxs-lookup"><span data-stu-id="8639f-1321">Removes an event handler for a</span></span>

<span data-ttu-id="8639f-1322">現在、サポートされているイベントの種類は、`Office.EventType.AppointmentTimeChanged`と`Office.EventType.RecipientsChanged`です。 `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="8639f-1322">Currently, the supported event types are `Office.EventType.AppointmentTimeChanged` and `Office.EventType.RecipientsChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8639f-1323">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="8639f-1323">Parameters:</span></span>

| <span data-ttu-id="8639f-1324">名前</span><span class="sxs-lookup"><span data-stu-id="8639f-1324">Name</span></span> | <span data-ttu-id="8639f-1325">種類</span><span class="sxs-lookup"><span data-stu-id="8639f-1325">Type</span></span> | <span data-ttu-id="8639f-1326">属性</span><span class="sxs-lookup"><span data-stu-id="8639f-1326">Attributes</span></span> | <span data-ttu-id="8639f-1327">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-1327">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="8639f-1328">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="8639f-1328">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="8639f-1329">ハンドラを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="8639f-1329">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="8639f-1330">関数</span><span class="sxs-lookup"><span data-stu-id="8639f-1330">Function</span></span> || <span data-ttu-id="8639f-p180">イベントを処理する関数。この関数は、オブジェクト リテラルである単一パラメータを受け入れる必要があります。パラメータの `type` プロパティは、`removeHandlerAsync` に渡される `eventType` パラメータと一致します。</span><span class="sxs-lookup"><span data-stu-id="8639f-p180">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="8639f-1334">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-1334">Object</span></span> | <span data-ttu-id="8639f-1335">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-1335">&lt;optional&gt;</span></span> | <span data-ttu-id="8639f-1336">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="8639f-1336">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="8639f-1337">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-1337">Object</span></span> | <span data-ttu-id="8639f-1338">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-1338">&lt;optional&gt;</span></span> | <span data-ttu-id="8639f-1339">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1339">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="8639f-1340">関数</span><span class="sxs-lookup"><span data-stu-id="8639f-1340">function</span></span>| <span data-ttu-id="8639f-1341">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-1341">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-1342">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="8639f-1342">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8639f-1343">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1343">Requirements</span></span>

|<span data-ttu-id="8639f-1344">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1344">Requirement</span></span>| <span data-ttu-id="8639f-1345">値</span><span class="sxs-lookup"><span data-stu-id="8639f-1345">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-1346">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-1346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8639f-1347">1.7</span><span class="sxs-lookup"><span data-stu-id="8639f-1347">-17</span></span> |
|[<span data-ttu-id="8639f-1348">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-1348">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8639f-1349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8639f-1349">ReadItem</span></span> |
|[<span data-ttu-id="8639f-1350">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-1350">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8639f-1351">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8639f-1351">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="8639f-1352">saveAsync ([オプション] 、コールバック)</span><span class="sxs-lookup"><span data-stu-id="8639f-1352">saveAsync([options], callback)</span></span>

<span data-ttu-id="8639f-1353">アイテムを非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="8639f-1353">Asynchronously saves an item.</span></span>

<span data-ttu-id="8639f-p181">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッド経由でアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-p181">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="8639f-1357">アドインが、WS または REST API を使用しようとして`itemId`を取得するために、新規作成モードでアイテム上の`saveAsync`を呼び出す場合、Outlook キャッシュ モードでは、アイテムがサーバーと実際に同期するまでに時間がかかる場合があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="8639f-1357">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="8639f-1358">項目が同期されるまで、 `itemId` を使用すると、エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1358">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="8639f-p183">予定はドラフト状態にはならないため、作成モードで予定に`saveAsync`が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="8639f-1362">次のクライアントは、新規作成モードで予定上の `saveAsync` に対して様々なふるまいをします。</span><span class="sxs-lookup"><span data-stu-id="8639f-1362">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="8639f-1363">Mac Outlook は、作成モードの会議で`saveAsync`をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="8639f-1363">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="8639f-1364">Mac Outlookの会議場で  `saveAsync` を呼びだすと、エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1364">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="8639f-1365">作成モードの予定上で`saveAsync`が呼び出されると、Outlook on the web は常に、招待または更新を送信します。</span><span class="sxs-lookup"><span data-stu-id="8639f-1365">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8639f-1366">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="8639f-1366">Parameters:</span></span>

|<span data-ttu-id="8639f-1367">名前</span><span class="sxs-lookup"><span data-stu-id="8639f-1367">Name</span></span>|<span data-ttu-id="8639f-1368">種類</span><span class="sxs-lookup"><span data-stu-id="8639f-1368">Type</span></span>|<span data-ttu-id="8639f-1369">属性</span><span class="sxs-lookup"><span data-stu-id="8639f-1369">Attributes</span></span>|<span data-ttu-id="8639f-1370">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-1370">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="8639f-1371">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-1371">Object</span></span>|<span data-ttu-id="8639f-1372">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-1372">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-1373">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="8639f-1373">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8639f-1374">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-1374">Object</span></span>|<span data-ttu-id="8639f-1375">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-1375">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-1376">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1376">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8639f-1377">関数</span><span class="sxs-lookup"><span data-stu-id="8639f-1377">function</span></span>||<span data-ttu-id="8639f-1378">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="8639f-1378">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8639f-1379">成功すると、アイテム識別子が`asyncResult.value`プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1379">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8639f-1380">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1380">Requirements</span></span>

|<span data-ttu-id="8639f-1381">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1381">Requirement</span></span>|<span data-ttu-id="8639f-1382">値</span><span class="sxs-lookup"><span data-stu-id="8639f-1382">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-1383">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-1383">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-1384">1.3</span><span class="sxs-lookup"><span data-stu-id="8639f-1384">1.3</span></span>|
|[<span data-ttu-id="8639f-1385">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-1385">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-1386">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8639f-1386">ReadWriteItem</span></span>|
|[<span data-ttu-id="8639f-1387">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-1387">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-1388">新規作成</span><span class="sxs-lookup"><span data-stu-id="8639f-1388">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8639f-1389">例</span><span class="sxs-lookup"><span data-stu-id="8639f-1389">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="8639f-p185">次の例は、コールバック関数に渡される `result` パラメータの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="8639f-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="8639f-1392">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="8639f-1392">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="8639f-1393">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="8639f-1393">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="8639f-p186">`setSelectedDataAsync`メソッドは、指定された文字列を項目のサブジェクトまたは本文のカーソル位置に挿入するか、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="8639f-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8639f-1397">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="8639f-1397">Parameters:</span></span>

|<span data-ttu-id="8639f-1398">名前</span><span class="sxs-lookup"><span data-stu-id="8639f-1398">Name</span></span>|<span data-ttu-id="8639f-1399">種類</span><span class="sxs-lookup"><span data-stu-id="8639f-1399">Type</span></span>|<span data-ttu-id="8639f-1400">属性</span><span class="sxs-lookup"><span data-stu-id="8639f-1400">Attributes</span></span>|<span data-ttu-id="8639f-1401">説明</span><span class="sxs-lookup"><span data-stu-id="8639f-1401">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="8639f-1402">文字列</span><span class="sxs-lookup"><span data-stu-id="8639f-1402">String</span></span>||<span data-ttu-id="8639f-p187">挿入されるデータです。データの長さは 1,000,000 文字以内です。1,000,000 文字を超えるデータが渡されると、 `ArgumentOutOfRange` の例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="8639f-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="8639f-1406">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-1406">Object</span></span>|<span data-ttu-id="8639f-1407">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-1407">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-1408">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="8639f-1408">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8639f-1409">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8639f-1409">Object</span></span>|<span data-ttu-id="8639f-1410">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-1410">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-1411">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1411">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="8639f-1412">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8639f-1412">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="8639f-1413">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8639f-1413">&lt;optional&gt;</span></span>|<span data-ttu-id="8639f-p188">`text` の場合、Outlook Web アプリ と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-p188">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="8639f-p189">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-p189">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="8639f-1418">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="8639f-1418">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="8639f-1419">関数</span><span class="sxs-lookup"><span data-stu-id="8639f-1419">function</span></span>||<span data-ttu-id="8639f-1420">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="8639f-1420">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8639f-1421">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1421">Requirements</span></span>

|<span data-ttu-id="8639f-1422">要件</span><span class="sxs-lookup"><span data-stu-id="8639f-1422">Requirement</span></span>|<span data-ttu-id="8639f-1423">値</span><span class="sxs-lookup"><span data-stu-id="8639f-1423">Value</span></span>|
|---|---|
|[<span data-ttu-id="8639f-1424">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8639f-1424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8639f-1425">1.2</span><span class="sxs-lookup"><span data-stu-id="8639f-1425">1.2</span></span>|
|[<span data-ttu-id="8639f-1426">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="8639f-1426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8639f-1427">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8639f-1427">ReadWriteItem</span></span>|
|[<span data-ttu-id="8639f-1428">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8639f-1428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8639f-1429">新規作成</span><span class="sxs-lookup"><span data-stu-id="8639f-1429">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8639f-1430">例</span><span class="sxs-lookup"><span data-stu-id="8639f-1430">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```