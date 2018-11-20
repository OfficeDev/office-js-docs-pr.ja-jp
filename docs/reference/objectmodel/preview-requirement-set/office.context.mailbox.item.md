
# <a name="item"></a><span data-ttu-id="162c4-101">item</span><span class="sxs-lookup"><span data-stu-id="162c4-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="162c4-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="162c4-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="162c4-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="162c4-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-105">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-105">Requirements</span></span>

|<span data-ttu-id="162c4-106">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-106">Requirement</span></span>|<span data-ttu-id="162c4-107">値</span><span class="sxs-lookup"><span data-stu-id="162c4-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-109">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-109">1.0</span></span>|
|[<span data-ttu-id="162c4-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="162c4-111">Restricted</span></span>|
|[<span data-ttu-id="162c4-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-113">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="162c4-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="162c4-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-114">Members and methods</span></span>

| <span data-ttu-id="162c4-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-115">Member</span></span> | <span data-ttu-id="162c4-116">種類</span><span class="sxs-lookup"><span data-stu-id="162c4-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="162c4-117">attachments</span><span class="sxs-lookup"><span data-stu-id="162c4-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="162c4-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-118">Member</span></span> |
| [<span data-ttu-id="162c4-119">bcc</span><span class="sxs-lookup"><span data-stu-id="162c4-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="162c4-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-120">Member</span></span> |
| [<span data-ttu-id="162c4-121">body</span><span class="sxs-lookup"><span data-stu-id="162c4-121">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="162c4-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-122">Member</span></span> |
| [<span data-ttu-id="162c4-123">cc</span><span class="sxs-lookup"><span data-stu-id="162c4-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="162c4-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-124">Member</span></span> |
| [<span data-ttu-id="162c4-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="162c4-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="162c4-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-126">Member</span></span> |
| [<span data-ttu-id="162c4-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="162c4-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="162c4-128">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-128">Member</span></span> |
| [<span data-ttu-id="162c4-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="162c4-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="162c4-130">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-130">Member</span></span> |
| [<span data-ttu-id="162c4-131">end</span><span class="sxs-lookup"><span data-stu-id="162c4-131">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="162c4-132">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-132">Member</span></span> |
| [<span data-ttu-id="162c4-133">from</span><span class="sxs-lookup"><span data-stu-id="162c4-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="162c4-134">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-134">Member</span></span> |
| [<span data-ttu-id="162c4-135">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="162c4-135">internetHeaders</span></span>](#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders) | <span data-ttu-id="162c4-136">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-136">Member</span></span> |
| [<span data-ttu-id="162c4-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="162c4-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="162c4-138">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-138">Member</span></span> |
| [<span data-ttu-id="162c4-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="162c4-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="162c4-140">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-140">Member</span></span> |
| [<span data-ttu-id="162c4-141">itemId</span><span class="sxs-lookup"><span data-stu-id="162c4-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="162c4-142">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-142">Member</span></span> |
| [<span data-ttu-id="162c4-143">itemType</span><span class="sxs-lookup"><span data-stu-id="162c4-143">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="162c4-144">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-144">Member</span></span> |
| [<span data-ttu-id="162c4-145">location</span><span class="sxs-lookup"><span data-stu-id="162c4-145">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="162c4-146">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-146">Member</span></span> |
| [<span data-ttu-id="162c4-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="162c4-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="162c4-148">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-148">Member</span></span> |
| [<span data-ttu-id="162c4-149">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="162c4-149">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="162c4-150">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-150">Member</span></span> |
| [<span data-ttu-id="162c4-151">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="162c4-151">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="162c4-152">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-152">Member</span></span> |
| [<span data-ttu-id="162c4-153">organizer</span><span class="sxs-lookup"><span data-stu-id="162c4-153">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="162c4-154">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-154">Member</span></span> |
| [<span data-ttu-id="162c4-155">recurrence</span><span class="sxs-lookup"><span data-stu-id="162c4-155">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="162c4-156">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-156">Member</span></span> |
| [<span data-ttu-id="162c4-157">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="162c4-157">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="162c4-158">Member</span><span class="sxs-lookup"><span data-stu-id="162c4-158">Member</span></span> |
| [<span data-ttu-id="162c4-159">sender</span><span class="sxs-lookup"><span data-stu-id="162c4-159">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="162c4-160">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-160">Member</span></span> |
| [<span data-ttu-id="162c4-161">seriesId</span><span class="sxs-lookup"><span data-stu-id="162c4-161">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="162c4-162">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-162">Member</span></span> |
| [<span data-ttu-id="162c4-163">start</span><span class="sxs-lookup"><span data-stu-id="162c4-163">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="162c4-164">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-164">Member</span></span> |
| [<span data-ttu-id="162c4-165">subject</span><span class="sxs-lookup"><span data-stu-id="162c4-165">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="162c4-166">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-166">Member</span></span> |
| [<span data-ttu-id="162c4-167">to</span><span class="sxs-lookup"><span data-stu-id="162c4-167">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="162c4-168">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-168">Member</span></span> |
| [<span data-ttu-id="162c4-169">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="162c4-169">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="162c4-170">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-170">Method</span></span> |
| [<span data-ttu-id="162c4-171">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="162c4-171">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="162c4-172">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-172">Method</span></span> |
| [<span data-ttu-id="162c4-173">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="162c4-173">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="162c4-174">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-174">Method</span></span> |
| [<span data-ttu-id="162c4-175">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="162c4-175">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="162c4-176">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-176">Method</span></span> |
| [<span data-ttu-id="162c4-177">close</span><span class="sxs-lookup"><span data-stu-id="162c4-177">close</span></span>](#close) | <span data-ttu-id="162c4-178">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-178">Method</span></span> |
| [<span data-ttu-id="162c4-179">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="162c4-179">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="162c4-180">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-180">Method</span></span> |
| [<span data-ttu-id="162c4-181">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="162c4-181">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="162c4-182">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-182">Method</span></span> |
| [<span data-ttu-id="162c4-183">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="162c4-183">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) | <span data-ttu-id="162c4-184">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-184">Method</span></span> |
| [<span data-ttu-id="162c4-185">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="162c4-185">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="162c4-186">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-186">Method</span></span> |
| [<span data-ttu-id="162c4-187">getEntities</span><span class="sxs-lookup"><span data-stu-id="162c4-187">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="162c4-188">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-188">Method</span></span> |
| [<span data-ttu-id="162c4-189">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="162c4-189">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="162c4-190">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-190">Method</span></span> |
| [<span data-ttu-id="162c4-191">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="162c4-191">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="162c4-192">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-192">Method</span></span> |
| [<span data-ttu-id="162c4-193">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="162c4-193">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="162c4-194">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-194">Method</span></span> |
| [<span data-ttu-id="162c4-195">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="162c4-195">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="162c4-196">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-196">Method</span></span> |
| [<span data-ttu-id="162c4-197">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="162c4-197">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="162c4-198">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-198">Method</span></span> |
| [<span data-ttu-id="162c4-199">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="162c4-199">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="162c4-200">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-200">Method</span></span> |
| [<span data-ttu-id="162c4-201">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="162c4-201">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="162c4-202">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-202">Method</span></span> |
| [<span data-ttu-id="162c4-203">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="162c4-203">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="162c4-204">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-204">Method</span></span> |
| [<span data-ttu-id="162c4-205">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="162c4-205">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="162c4-206">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-206">Method</span></span> |
| [<span data-ttu-id="162c4-207">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="162c4-207">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="162c4-208">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-208">Method</span></span> |
| [<span data-ttu-id="162c4-209">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="162c4-209">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="162c4-210">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-210">Method</span></span> |
| [<span data-ttu-id="162c4-211">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="162c4-211">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="162c4-212">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-212">Method</span></span> |
| [<span data-ttu-id="162c4-213">saveAsync</span><span class="sxs-lookup"><span data-stu-id="162c4-213">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="162c4-214">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-214">Method</span></span> |
| [<span data-ttu-id="162c4-215">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="162c4-215">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="162c4-216">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-216">Method</span></span> |

### <a name="example"></a><span data-ttu-id="162c4-217">例</span><span class="sxs-lookup"><span data-stu-id="162c4-217">Example</span></span>

<span data-ttu-id="162c4-218">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="162c4-218">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
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

### <a name="members"></a><span data-ttu-id="162c4-219">メンバー</span><span class="sxs-lookup"><span data-stu-id="162c4-219">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="162c4-220">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="162c4-220">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="162c4-221">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="162c4-221">Gets the item's attachments as an array.</span></span> <span data-ttu-id="162c4-222">閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="162c4-222">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="162c4-223">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="162c4-223">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="162c4-224">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="162c4-224">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-225">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-225">Type:</span></span>

*   <span data-ttu-id="162c4-226">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="162c4-226">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-227">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-227">Requirements</span></span>

|<span data-ttu-id="162c4-228">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-228">Requirement</span></span>|<span data-ttu-id="162c4-229">値</span><span class="sxs-lookup"><span data-stu-id="162c4-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-230">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-230">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-231">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-231">1.0</span></span>|
|[<span data-ttu-id="162c4-232">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-232">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-233">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-233">ReadItem</span></span>|
|[<span data-ttu-id="162c4-234">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-234">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-235">Read</span><span class="sxs-lookup"><span data-stu-id="162c4-235">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-236">例</span><span class="sxs-lookup"><span data-stu-id="162c4-236">Example</span></span>

<span data-ttu-id="162c4-237">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="162c4-237">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="162c4-238">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="162c4-238">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="162c4-239">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="162c4-239">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="162c4-240">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="162c4-240">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-241">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-241">Type:</span></span>

*   [<span data-ttu-id="162c4-242">Recipients</span><span class="sxs-lookup"><span data-stu-id="162c4-242">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="162c4-243">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-243">Requirements</span></span>

|<span data-ttu-id="162c4-244">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-244">Requirement</span></span>|<span data-ttu-id="162c4-245">値</span><span class="sxs-lookup"><span data-stu-id="162c4-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-246">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-246">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-247">1.1</span><span class="sxs-lookup"><span data-stu-id="162c4-247">1.1</span></span>|
|[<span data-ttu-id="162c4-248">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-249">ReadItem</span></span>|
|[<span data-ttu-id="162c4-250">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-251">Compose</span><span class="sxs-lookup"><span data-stu-id="162c4-251">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-252">例</span><span class="sxs-lookup"><span data-stu-id="162c4-252">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="162c4-253">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="162c4-253">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="162c4-254">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="162c4-254">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-255">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-255">Type:</span></span>

*   [<span data-ttu-id="162c4-256">Body</span><span class="sxs-lookup"><span data-stu-id="162c4-256">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="162c4-257">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-257">Requirements</span></span>

|<span data-ttu-id="162c4-258">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-258">Requirement</span></span>|<span data-ttu-id="162c4-259">値</span><span class="sxs-lookup"><span data-stu-id="162c4-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-260">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-260">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-261">1.1</span><span class="sxs-lookup"><span data-stu-id="162c4-261">1.1</span></span>|
|[<span data-ttu-id="162c4-262">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-262">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-263">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-263">ReadItem</span></span>|
|[<span data-ttu-id="162c4-264">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-264">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-265">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="162c4-265">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="162c4-266">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="162c4-266">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="162c4-267">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="162c4-267">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="162c4-268">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="162c4-268">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="162c4-269">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="162c4-269">Read mode</span></span>

<span data-ttu-id="162c4-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="162c4-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="162c4-272">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="162c4-272">Compose mode</span></span>

<span data-ttu-id="162c4-273">`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-273">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-274">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-274">Type:</span></span>

*   <span data-ttu-id="162c4-275">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="162c4-275">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-276">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-276">Requirements</span></span>

|<span data-ttu-id="162c4-277">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-277">Requirement</span></span>|<span data-ttu-id="162c4-278">値</span><span class="sxs-lookup"><span data-stu-id="162c4-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-279">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-280">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-280">1.0</span></span>|
|[<span data-ttu-id="162c4-281">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-282">ReadItem</span></span>|
|[<span data-ttu-id="162c4-283">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-284">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="162c4-284">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-285">例</span><span class="sxs-lookup"><span data-stu-id="162c4-285">Example</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="162c4-286">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="162c4-286">(nullable) conversationId :String</span></span>

<span data-ttu-id="162c4-287">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="162c4-287">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="162c4-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="162c4-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="162c4-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-292">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-292">Type:</span></span>

*   <span data-ttu-id="162c4-293">String</span><span class="sxs-lookup"><span data-stu-id="162c4-293">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-294">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-294">Requirements</span></span>

|<span data-ttu-id="162c4-295">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-295">Requirement</span></span>|<span data-ttu-id="162c4-296">値</span><span class="sxs-lookup"><span data-stu-id="162c4-296">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-297">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-297">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-298">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-298">1.0</span></span>|
|[<span data-ttu-id="162c4-299">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-299">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-300">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-300">ReadItem</span></span>|
|[<span data-ttu-id="162c4-301">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-301">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-302">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="162c4-302">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="162c4-303">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="162c4-303">dateTimeCreated :Date</span></span>

<span data-ttu-id="162c4-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="162c4-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-306">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-306">Type:</span></span>

*   <span data-ttu-id="162c4-307">日付</span><span class="sxs-lookup"><span data-stu-id="162c4-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-308">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-308">Requirements</span></span>

|<span data-ttu-id="162c4-309">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-309">Requirement</span></span>|<span data-ttu-id="162c4-310">値</span><span class="sxs-lookup"><span data-stu-id="162c4-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-312">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-312">1.0</span></span>|
|[<span data-ttu-id="162c4-313">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-313">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-314">ReadItem</span></span>|
|[<span data-ttu-id="162c4-315">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-315">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-316">Read</span><span class="sxs-lookup"><span data-stu-id="162c4-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-317">例</span><span class="sxs-lookup"><span data-stu-id="162c4-317">Example</span></span>

```javascript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="162c4-318">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="162c4-318">dateTimeModified :Date</span></span>

<span data-ttu-id="162c4-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="162c4-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="162c4-321">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="162c4-321">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-322">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-322">Type:</span></span>

*   <span data-ttu-id="162c4-323">日付</span><span class="sxs-lookup"><span data-stu-id="162c4-323">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-324">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-324">Requirements</span></span>

|<span data-ttu-id="162c4-325">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-325">Requirement</span></span>|<span data-ttu-id="162c4-326">値</span><span class="sxs-lookup"><span data-stu-id="162c4-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-327">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-328">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-328">1.0</span></span>|
|[<span data-ttu-id="162c4-329">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-329">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-330">ReadItem</span></span>|
|[<span data-ttu-id="162c4-331">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-331">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-332">Read</span><span class="sxs-lookup"><span data-stu-id="162c4-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-333">例</span><span class="sxs-lookup"><span data-stu-id="162c4-333">Example</span></span>

```javascript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="162c4-334">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="162c4-334">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="162c4-335">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="162c4-335">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="162c4-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="162c4-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="162c4-338">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="162c4-338">Read mode</span></span>

<span data-ttu-id="162c4-339">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-339">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="162c4-340">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="162c4-340">Compose mode</span></span>

<span data-ttu-id="162c4-341">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-341">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="162c4-342">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="162c4-342">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-343">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-343">Type:</span></span>

*   <span data-ttu-id="162c4-344">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="162c4-344">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-345">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-345">Requirements</span></span>

|<span data-ttu-id="162c4-346">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-346">Requirement</span></span>|<span data-ttu-id="162c4-347">値</span><span class="sxs-lookup"><span data-stu-id="162c4-347">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-348">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-349">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-349">1.0</span></span>|
|[<span data-ttu-id="162c4-350">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-350">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-351">ReadItem</span></span>|
|[<span data-ttu-id="162c4-352">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-352">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-353">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="162c4-353">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-354">例</span><span class="sxs-lookup"><span data-stu-id="162c4-354">Example</span></span>

<span data-ttu-id="162c4-355">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="162c4-355">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="162c4-356">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="162c4-356">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="162c4-357">メッセージの送信者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="162c4-357">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="162c4-p112">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="162c4-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="162c4-360">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="162c4-360">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="162c4-361">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="162c4-361">Read mode</span></span>

<span data-ttu-id="162c4-362">`from` プロパティは `EmailAddressDetails` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-362">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="162c4-363">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="162c4-363">Compose mode</span></span>

<span data-ttu-id="162c4-364">`from` プロパティは From 値を取得するメソッドを提供する `From` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-364">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="162c4-365">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-365">Type:</span></span>

*   <span data-ttu-id="162c4-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="162c4-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-367">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-367">Requirements</span></span>

|<span data-ttu-id="162c4-368">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-368">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="162c4-369">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-369">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-370">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-370">1.0</span></span>|<span data-ttu-id="162c4-371">1.7</span><span class="sxs-lookup"><span data-stu-id="162c4-371">-17</span></span>|
|[<span data-ttu-id="162c4-372">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-372">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-373">ReadItem</span></span>|<span data-ttu-id="162c4-374">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="162c4-374">ReadWriteItem</span></span>|
|[<span data-ttu-id="162c4-375">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-375">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-376">Read</span><span class="sxs-lookup"><span data-stu-id="162c4-376">Read</span></span>|<span data-ttu-id="162c4-377">Compose</span><span class="sxs-lookup"><span data-stu-id="162c4-377">Compose</span></span>|

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="162c4-378">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="162c4-378">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="162c4-379">メッセージのインターネット ヘッダーを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="162c4-379">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-380">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-380">Type:</span></span>

*   [<span data-ttu-id="162c4-381">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="162c4-381">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="162c4-382">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-382">Requirements</span></span>

|<span data-ttu-id="162c4-383">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-383">Requirement</span></span>|<span data-ttu-id="162c4-384">値</span><span class="sxs-lookup"><span data-stu-id="162c4-384">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-385">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-385">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-386">プレビュー</span><span class="sxs-lookup"><span data-stu-id="162c4-386">Preview</span></span>|
|[<span data-ttu-id="162c4-387">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-387">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-388">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-388">ReadItem</span></span>|
|[<span data-ttu-id="162c4-389">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-389">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-390">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="162c4-390">Compose or read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="162c4-391">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="162c4-391">internetMessageId :String</span></span>

<span data-ttu-id="162c4-p113">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="162c4-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-394">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-394">Type:</span></span>

*   <span data-ttu-id="162c4-395">String</span><span class="sxs-lookup"><span data-stu-id="162c4-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-396">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-396">Requirements</span></span>

|<span data-ttu-id="162c4-397">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-397">Requirement</span></span>|<span data-ttu-id="162c4-398">値</span><span class="sxs-lookup"><span data-stu-id="162c4-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-399">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-399">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-400">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-400">1.0</span></span>|
|[<span data-ttu-id="162c4-401">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-401">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-402">ReadItem</span></span>|
|[<span data-ttu-id="162c4-403">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-403">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-404">Read</span><span class="sxs-lookup"><span data-stu-id="162c4-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-405">例</span><span class="sxs-lookup"><span data-stu-id="162c4-405">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="162c4-406">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="162c4-406">itemClass :String</span></span>

<span data-ttu-id="162c4-p114">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="162c4-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="162c4-p115">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="162c4-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="162c4-411">型</span><span class="sxs-lookup"><span data-stu-id="162c4-411">Type</span></span>|<span data-ttu-id="162c4-412">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-412">Description</span></span>|<span data-ttu-id="162c4-413">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="162c4-413">item class</span></span>|
|---|---|---|
|<span data-ttu-id="162c4-414">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="162c4-414">Appointment items</span></span>|<span data-ttu-id="162c4-415">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="162c4-415">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="162c4-416">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="162c4-416">Message items</span></span>|<span data-ttu-id="162c4-417">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="162c4-417">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="162c4-418">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="162c4-418">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-419">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-419">Type:</span></span>

*   <span data-ttu-id="162c4-420">String</span><span class="sxs-lookup"><span data-stu-id="162c4-420">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-421">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-421">Requirements</span></span>

|<span data-ttu-id="162c4-422">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-422">Requirement</span></span>|<span data-ttu-id="162c4-423">値</span><span class="sxs-lookup"><span data-stu-id="162c4-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-424">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-425">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-425">1.0</span></span>|
|[<span data-ttu-id="162c4-426">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-427">ReadItem</span></span>|
|[<span data-ttu-id="162c4-428">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-429">Read</span><span class="sxs-lookup"><span data-stu-id="162c4-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-430">例</span><span class="sxs-lookup"><span data-stu-id="162c4-430">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="162c4-431">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="162c4-431">(nullable) itemId :String</span></span>

<span data-ttu-id="162c4-p116">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="162c4-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="162c4-434">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="162c4-434">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="162c4-435">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="162c4-435">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="162c4-436">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="162c4-436">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="162c4-437">詳細については、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="162c4-437">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="162c4-p118">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-440">種類:</span><span class="sxs-lookup"><span data-stu-id="162c4-440">Type:</span></span>

*   <span data-ttu-id="162c4-441">String</span><span class="sxs-lookup"><span data-stu-id="162c4-441">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-442">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-442">Requirements</span></span>

|<span data-ttu-id="162c4-443">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-443">Requirement</span></span>|<span data-ttu-id="162c4-444">値</span><span class="sxs-lookup"><span data-stu-id="162c4-444">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-445">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-445">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-446">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-446">1.0</span></span>|
|[<span data-ttu-id="162c4-447">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-447">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-448">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-448">ReadItem</span></span>|
|[<span data-ttu-id="162c4-449">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-449">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-450">Read</span><span class="sxs-lookup"><span data-stu-id="162c4-450">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-451">例</span><span class="sxs-lookup"><span data-stu-id="162c4-451">Example</span></span>

<span data-ttu-id="162c4-p119">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="162c4-454">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="162c4-454">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="162c4-455">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="162c4-455">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="162c4-456">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="162c4-456">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-457">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-457">Type:</span></span>

*   [<span data-ttu-id="162c4-458">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="162c4-458">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="162c4-459">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-459">Requirements</span></span>

|<span data-ttu-id="162c4-460">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-460">Requirement</span></span>|<span data-ttu-id="162c4-461">値</span><span class="sxs-lookup"><span data-stu-id="162c4-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-462">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-463">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-463">1.0</span></span>|
|[<span data-ttu-id="162c4-464">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-464">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-465">ReadItem</span></span>|
|[<span data-ttu-id="162c4-466">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-466">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-467">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="162c4-467">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-468">例</span><span class="sxs-lookup"><span data-stu-id="162c4-468">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="162c4-469">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="162c4-469">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="162c4-470">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="162c4-470">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="162c4-471">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="162c4-471">Read mode</span></span>

<span data-ttu-id="162c4-472">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-472">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="162c4-473">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="162c4-473">Compose mode</span></span>

<span data-ttu-id="162c4-474">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-474">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-475">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-475">Type:</span></span>

*   <span data-ttu-id="162c4-476">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="162c4-476">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-477">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-477">Requirements</span></span>

|<span data-ttu-id="162c4-478">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-478">Requirement</span></span>|<span data-ttu-id="162c4-479">値</span><span class="sxs-lookup"><span data-stu-id="162c4-479">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-480">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-480">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-481">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-481">1.0</span></span>|
|[<span data-ttu-id="162c4-482">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-482">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-483">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-483">ReadItem</span></span>|
|[<span data-ttu-id="162c4-484">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-484">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-485">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="162c4-485">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-486">例</span><span class="sxs-lookup"><span data-stu-id="162c4-486">Example</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="162c4-487">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="162c4-487">normalizedSubject :String</span></span>

<span data-ttu-id="162c4-p120">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="162c4-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="162c4-p121">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="162c4-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-492">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-492">Type:</span></span>

*   <span data-ttu-id="162c4-493">String</span><span class="sxs-lookup"><span data-stu-id="162c4-493">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-494">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-494">Requirements</span></span>

|<span data-ttu-id="162c4-495">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-495">Requirement</span></span>|<span data-ttu-id="162c4-496">値</span><span class="sxs-lookup"><span data-stu-id="162c4-496">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-497">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-497">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-498">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-498">1.0</span></span>|
|[<span data-ttu-id="162c4-499">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-499">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-500">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-500">ReadItem</span></span>|
|[<span data-ttu-id="162c4-501">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-501">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-502">Read</span><span class="sxs-lookup"><span data-stu-id="162c4-502">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-503">例</span><span class="sxs-lookup"><span data-stu-id="162c4-503">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="162c4-504">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="162c4-504">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="162c4-505">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="162c4-505">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-506">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-506">Type:</span></span>

*   [<span data-ttu-id="162c4-507">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="162c4-507">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="162c4-508">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-508">Requirements</span></span>

|<span data-ttu-id="162c4-509">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-509">Requirement</span></span>|<span data-ttu-id="162c4-510">値</span><span class="sxs-lookup"><span data-stu-id="162c4-510">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-511">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-511">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-512">1.3</span><span class="sxs-lookup"><span data-stu-id="162c4-512">1.3</span></span>|
|[<span data-ttu-id="162c4-513">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-513">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-514">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-514">ReadItem</span></span>|
|[<span data-ttu-id="162c4-515">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-515">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-516">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="162c4-516">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="162c4-517">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="162c4-517">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="162c4-518">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="162c4-518">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="162c4-519">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="162c4-519">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="162c4-520">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="162c4-520">Read mode</span></span>

<span data-ttu-id="162c4-521">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-521">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="162c4-522">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="162c4-522">Compose mode</span></span>

<span data-ttu-id="162c4-523">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-523">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-524">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-524">Type:</span></span>

*   <span data-ttu-id="162c4-525">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="162c4-525">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-526">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-526">Requirements</span></span>

|<span data-ttu-id="162c4-527">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-527">Requirement</span></span>|<span data-ttu-id="162c4-528">値</span><span class="sxs-lookup"><span data-stu-id="162c4-528">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-529">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-529">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-530">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-530">1.0</span></span>|
|[<span data-ttu-id="162c4-531">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-531">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-532">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-532">ReadItem</span></span>|
|[<span data-ttu-id="162c4-533">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-533">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-534">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="162c4-534">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-535">例</span><span class="sxs-lookup"><span data-stu-id="162c4-535">Example</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="162c4-536">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="162c4-536">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="162c4-537">指定の会議の開催者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="162c4-537">Gets the email address of the meeting organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="162c4-538">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="162c4-538">Read mode</span></span>

<span data-ttu-id="162c4-539">`organizer` プロパティは、会議開催者を表す [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-539">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="162c4-540">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="162c4-540">Compose mode</span></span>

<span data-ttu-id="162c4-541">`organizer` プロパティは Organizer 値を取得するメソッドを提供する [Organizer](/javascript/api/outlook/office.organizer) オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-541">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-542">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-542">Type:</span></span>

*   <span data-ttu-id="162c4-543">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="162c4-543">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-544">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-544">Requirements</span></span>

|<span data-ttu-id="162c4-545">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-545">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="162c4-546">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-547">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-547">1.0</span></span>|<span data-ttu-id="162c4-548">1.7</span><span class="sxs-lookup"><span data-stu-id="162c4-548">-17</span></span>|
|[<span data-ttu-id="162c4-549">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-549">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-550">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-550">ReadItem</span></span>|<span data-ttu-id="162c4-551">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="162c4-551">ReadWriteItem</span></span>|
|[<span data-ttu-id="162c4-552">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-552">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-553">Read</span><span class="sxs-lookup"><span data-stu-id="162c4-553">Read</span></span>|<span data-ttu-id="162c4-554">Compose</span><span class="sxs-lookup"><span data-stu-id="162c4-554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-555">例</span><span class="sxs-lookup"><span data-stu-id="162c4-555">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="162c4-556">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="162c4-556">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="162c4-557">予定の定期的なパターンを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="162c4-557">Gets or sets the location of an appointment.</span></span> <span data-ttu-id="162c4-558">会議出席依頼の定期的なパターンを取得します。</span><span class="sxs-lookup"><span data-stu-id="162c4-558">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="162c4-559">予定アイテムの閲覧モードと新規作成モード。</span><span class="sxs-lookup"><span data-stu-id="162c4-559">Read and compose modes for appointment items.</span></span> <span data-ttu-id="162c4-560">会議出席依頼アイテムの閲覧モード。</span><span class="sxs-lookup"><span data-stu-id="162c4-560">Read mode for meeting request items.</span></span>

<span data-ttu-id="162c4-561">`recurrence` プロパティは、アイテムがシリーズか、シリーズに含まれるインスタンスの場合、定期的な予定または会議出席依頼に対して [recurrence](/javascript/api/outlook/office.recurrence) オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-561">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="162c4-562">`null` は、単発の予定および単発の予定の会議出席依頼に対して返されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-562">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="162c4-563">`undefined` は、会議出席依頼ではないメッセージに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-563">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="162c4-564">注: 会議出席依頼の `itemClass` 値は IPM.Schedule.Meeting.Request です。</span><span class="sxs-lookup"><span data-stu-id="162c4-564">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="162c4-565">注: recurrence オブジェクトが `null` の場合、オブジェクトがシリーズの一部ではなく、1 つの単発の予定または 1 つの単発の予定の会議出席依頼であることを示します。</span><span class="sxs-lookup"><span data-stu-id="162c4-565">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-566">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-566">Type:</span></span>

* [<span data-ttu-id="162c4-567">Recurrence</span><span class="sxs-lookup"><span data-stu-id="162c4-567">recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="162c4-568">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-568">Requirement</span></span>|<span data-ttu-id="162c4-569">値</span><span class="sxs-lookup"><span data-stu-id="162c4-569">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-570">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-571">1.7</span><span class="sxs-lookup"><span data-stu-id="162c4-571">-17</span></span>|
|[<span data-ttu-id="162c4-572">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-572">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-573">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-573">ReadItem</span></span>|
|[<span data-ttu-id="162c4-574">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-574">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-575">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="162c4-575">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="162c4-576">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="162c4-576">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="162c4-577">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="162c4-577">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="162c4-578">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="162c4-578">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="162c4-579">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="162c4-579">Read mode</span></span>

<span data-ttu-id="162c4-580">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-580">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="162c4-581">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="162c4-581">Compose mode</span></span>

<span data-ttu-id="162c4-582">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-582">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-583">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-583">Type:</span></span>

*   <span data-ttu-id="162c4-584">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="162c4-584">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-585">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-585">Requirements</span></span>

|<span data-ttu-id="162c4-586">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-586">Requirement</span></span>|<span data-ttu-id="162c4-587">値</span><span class="sxs-lookup"><span data-stu-id="162c4-587">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-588">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-588">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-589">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-589">1.0</span></span>|
|[<span data-ttu-id="162c4-590">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-590">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-591">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-591">ReadItem</span></span>|
|[<span data-ttu-id="162c4-592">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-592">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-593">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="162c4-593">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-594">例</span><span class="sxs-lookup"><span data-stu-id="162c4-594">Example</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="162c4-595">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="162c4-595">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="162c4-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="162c4-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="162c4-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="162c4-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="162c4-600">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="162c4-600">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-601">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-601">Type:</span></span>

*   [<span data-ttu-id="162c4-602">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="162c4-602">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="162c4-603">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-603">Requirements</span></span>

|<span data-ttu-id="162c4-604">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-604">Requirement</span></span>|<span data-ttu-id="162c4-605">値</span><span class="sxs-lookup"><span data-stu-id="162c4-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-606">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-607">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-607">1.0</span></span>|
|[<span data-ttu-id="162c4-608">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-608">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-609">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-609">ReadItem</span></span>|
|[<span data-ttu-id="162c4-610">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-610">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-611">Read</span><span class="sxs-lookup"><span data-stu-id="162c4-611">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-612">例</span><span class="sxs-lookup"><span data-stu-id="162c4-612">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="162c4-613">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="162c4-613">(nullable) seriesId :String</span></span>

<span data-ttu-id="162c4-614">あるインスタンスが属するシリーズの ID を取得します。</span><span class="sxs-lookup"><span data-stu-id="162c4-614">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="162c4-615">OWA と Outlook では、`seriesId` はこのアイテムが属する親 (シリーズ) アイテムの Exchange Web Services (EWS) ID を返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-615">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="162c4-616">ただし、iOS と Android の場合、`seriesId` は親アイテムの REST ID を返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-616">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="162c4-617">`seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="162c4-617">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="162c4-618">`seriesId` プロパティは、Outlook REST API で使用される Outlook ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="162c4-618">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="162c4-619">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="162c4-619">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="162c4-620">詳細については、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="162c4-620">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="162c4-621">`seriesId` プロパティは、単発の予定、シリーズ アイテム、会議出席依頼など、親アイテムを持たないアイテムに対して `null` を返し、会議出席依頼ではないその他のアイテムに対して `undefined` を返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-621">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-622">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-622">Type:</span></span>

* <span data-ttu-id="162c4-623">String</span><span class="sxs-lookup"><span data-stu-id="162c4-623">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-624">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-624">Requirements</span></span>

|<span data-ttu-id="162c4-625">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-625">Requirement</span></span>|<span data-ttu-id="162c4-626">値</span><span class="sxs-lookup"><span data-stu-id="162c4-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-627">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-628">1.7</span><span class="sxs-lookup"><span data-stu-id="162c4-628">-17</span></span>|
|[<span data-ttu-id="162c4-629">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-629">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-630">ReadItem</span></span>|
|[<span data-ttu-id="162c4-631">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-631">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-632">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="162c4-632">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-633">例</span><span class="sxs-lookup"><span data-stu-id="162c4-633">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="162c4-634">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="162c4-634">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="162c4-635">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="162c4-635">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="162c4-p130">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="162c4-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="162c4-638">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="162c4-638">Read mode</span></span>

<span data-ttu-id="162c4-639">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-639">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="162c4-640">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="162c4-640">Compose mode</span></span>

<span data-ttu-id="162c4-641">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-641">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="162c4-642">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="162c4-642">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-643">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-643">Type:</span></span>

*   <span data-ttu-id="162c4-644">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="162c4-644">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-645">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-645">Requirements</span></span>

|<span data-ttu-id="162c4-646">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-646">Requirement</span></span>|<span data-ttu-id="162c4-647">値</span><span class="sxs-lookup"><span data-stu-id="162c4-647">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-648">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-648">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-649">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-649">1.0</span></span>|
|[<span data-ttu-id="162c4-650">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-650">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-651">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-651">ReadItem</span></span>|
|[<span data-ttu-id="162c4-652">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-652">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-653">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="162c4-653">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-654">例</span><span class="sxs-lookup"><span data-stu-id="162c4-654">Example</span></span>

<span data-ttu-id="162c4-655">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="162c4-655">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="162c4-656">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="162c4-656">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="162c4-657">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="162c4-657">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="162c4-658">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="162c4-658">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="162c4-659">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="162c4-659">Read mode</span></span>

<span data-ttu-id="162c4-p131">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="162c4-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="162c4-662">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="162c4-662">Compose mode</span></span>

<span data-ttu-id="162c4-663">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-663">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="162c4-664">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-664">Type:</span></span>

*   <span data-ttu-id="162c4-665">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="162c4-665">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-666">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-666">Requirements</span></span>

|<span data-ttu-id="162c4-667">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-667">Requirement</span></span>|<span data-ttu-id="162c4-668">値</span><span class="sxs-lookup"><span data-stu-id="162c4-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-669">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-670">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-670">1.0</span></span>|
|[<span data-ttu-id="162c4-671">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-671">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-672">ReadItem</span></span>|
|[<span data-ttu-id="162c4-673">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-673">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-674">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="162c4-674">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="162c4-675">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="162c4-675">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="162c4-676">メッセージの**宛先**行の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="162c4-676">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="162c4-677">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="162c4-677">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="162c4-678">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="162c4-678">Read mode</span></span>

<span data-ttu-id="162c4-p133">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="162c4-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="162c4-681">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="162c4-681">Compose mode</span></span>

<span data-ttu-id="162c4-682">`to` プロパティは、メッセージの**宛先**行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-682">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="162c4-683">型:</span><span class="sxs-lookup"><span data-stu-id="162c4-683">Type:</span></span>

*   <span data-ttu-id="162c4-684">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="162c4-684">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-685">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-685">Requirements</span></span>

|<span data-ttu-id="162c4-686">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-686">Requirement</span></span>|<span data-ttu-id="162c4-687">値</span><span class="sxs-lookup"><span data-stu-id="162c4-687">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-688">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-688">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-689">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-689">1.0</span></span>|
|[<span data-ttu-id="162c4-690">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-690">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-691">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-691">ReadItem</span></span>|
|[<span data-ttu-id="162c4-692">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-692">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-693">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="162c4-693">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-694">例</span><span class="sxs-lookup"><span data-stu-id="162c4-694">Example</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="162c4-695">メソッド</span><span class="sxs-lookup"><span data-stu-id="162c4-695">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="162c4-696">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="162c4-696">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="162c4-697">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="162c4-697">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="162c4-698">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="162c4-698">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="162c4-699">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="162c4-699">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="162c4-700">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="162c4-700">Parameters:</span></span>
|<span data-ttu-id="162c4-701">名前</span><span class="sxs-lookup"><span data-stu-id="162c4-701">Name</span></span>|<span data-ttu-id="162c4-702">型</span><span class="sxs-lookup"><span data-stu-id="162c4-702">Type</span></span>|<span data-ttu-id="162c4-703">属性</span><span class="sxs-lookup"><span data-stu-id="162c4-703">Attributes</span></span>|<span data-ttu-id="162c4-704">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-704">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="162c4-705">String</span><span class="sxs-lookup"><span data-stu-id="162c4-705">String</span></span>||<span data-ttu-id="162c4-p134">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="162c4-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="162c4-708">String</span><span class="sxs-lookup"><span data-stu-id="162c4-708">String</span></span>||<span data-ttu-id="162c4-p135">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="162c4-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="162c4-711">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-711">Object</span></span>|<span data-ttu-id="162c4-712">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-712">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-713">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="162c4-713">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="162c4-714">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-714">Object</span></span>|<span data-ttu-id="162c4-715">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-715">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-716">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="162c4-716">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="162c4-717">Boolean</span><span class="sxs-lookup"><span data-stu-id="162c4-717">Boolean</span></span>|<span data-ttu-id="162c4-718">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-718">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-719">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="162c4-719">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="162c4-720">function</span><span class="sxs-lookup"><span data-stu-id="162c4-720">function</span></span>|<span data-ttu-id="162c4-721">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-721">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-722">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-722">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="162c4-723">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-723">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="162c4-724">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="162c4-724">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="162c4-725">エラー</span><span class="sxs-lookup"><span data-stu-id="162c4-725">Errors</span></span>

|<span data-ttu-id="162c4-726">エラー コード</span><span class="sxs-lookup"><span data-stu-id="162c4-726">Error code</span></span>|<span data-ttu-id="162c4-727">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-727">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="162c4-728">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="162c4-728">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="162c4-729">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="162c4-729">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="162c4-730">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="162c4-730">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="162c4-731">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-731">Requirements</span></span>

|<span data-ttu-id="162c4-732">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-732">Requirement</span></span>|<span data-ttu-id="162c4-733">値</span><span class="sxs-lookup"><span data-stu-id="162c4-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-734">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-735">1.1</span><span class="sxs-lookup"><span data-stu-id="162c4-735">1.1</span></span>|
|[<span data-ttu-id="162c4-736">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-736">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-737">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="162c4-737">ReadWriteItem</span></span>|
|[<span data-ttu-id="162c4-738">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-738">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-739">Compose</span><span class="sxs-lookup"><span data-stu-id="162c4-739">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="162c4-740">例</span><span class="sxs-lookup"><span data-stu-id="162c4-740">Examples</span></span>

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

<span data-ttu-id="162c4-741">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="162c4-741">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="162c4-742">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="162c4-742">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="162c4-743">ファイルを添付ファイルとして base64 エンコーディングからメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="162c4-743">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="162c4-744">`addFileAttachmentFromBase64Async` メソッドは、base64 エンコーディングからファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="162c4-744">The `addFileAttachmentFromBase64Async` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span> <span data-ttu-id="162c4-745">このメソッドによって、AsyncResult.value オブジェクトの添付ファイル識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-745">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="162c4-746">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="162c4-746">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="162c4-747">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="162c4-747">Parameters:</span></span>
|<span data-ttu-id="162c4-748">名前</span><span class="sxs-lookup"><span data-stu-id="162c4-748">Name</span></span>|<span data-ttu-id="162c4-749">型</span><span class="sxs-lookup"><span data-stu-id="162c4-749">Type</span></span>|<span data-ttu-id="162c4-750">属性</span><span class="sxs-lookup"><span data-stu-id="162c4-750">Attributes</span></span>|<span data-ttu-id="162c4-751">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-751">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="162c4-752">String</span><span class="sxs-lookup"><span data-stu-id="162c4-752">String</span></span>||<span data-ttu-id="162c4-753">電子メールまたはイベントに追加する画像またはファイルの base64 でエンコードされたコンテンツ。</span><span class="sxs-lookup"><span data-stu-id="162c4-753">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="162c4-754">String</span><span class="sxs-lookup"><span data-stu-id="162c4-754">String</span></span>||<span data-ttu-id="162c4-p137">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="162c4-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="162c4-757">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-757">Object</span></span>|<span data-ttu-id="162c4-758">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-758">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-759">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="162c4-759">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="162c4-760">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-760">Object</span></span>|<span data-ttu-id="162c4-761">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-761">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-762">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="162c4-762">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="162c4-763">Boolean</span><span class="sxs-lookup"><span data-stu-id="162c4-763">Boolean</span></span>|<span data-ttu-id="162c4-764">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-764">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-765">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="162c4-765">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="162c4-766">function</span><span class="sxs-lookup"><span data-stu-id="162c4-766">function</span></span>|<span data-ttu-id="162c4-767">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-767">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-768">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-768">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="162c4-769">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-769">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="162c4-770">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="162c4-770">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="162c4-771">エラー</span><span class="sxs-lookup"><span data-stu-id="162c4-771">Errors</span></span>

|<span data-ttu-id="162c4-772">エラー コード</span><span class="sxs-lookup"><span data-stu-id="162c4-772">Error code</span></span>|<span data-ttu-id="162c4-773">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-773">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="162c4-774">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="162c4-774">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="162c4-775">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="162c4-775">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="162c4-776">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="162c4-776">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="162c4-777">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-777">Requirements</span></span>

|<span data-ttu-id="162c4-778">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-778">Requirement</span></span>|<span data-ttu-id="162c4-779">値</span><span class="sxs-lookup"><span data-stu-id="162c4-779">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-780">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-780">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-781">プレビュー</span><span class="sxs-lookup"><span data-stu-id="162c4-781">Preview</span></span>|
|[<span data-ttu-id="162c4-782">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-782">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-783">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="162c4-783">ReadWriteItem</span></span>|
|[<span data-ttu-id="162c4-784">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-784">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-785">Compose</span><span class="sxs-lookup"><span data-stu-id="162c4-785">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="162c4-786">例</span><span class="sxs-lookup"><span data-stu-id="162c4-786">Examples</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="162c4-787">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="162c4-787">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="162c4-788">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="162c4-788">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="162c4-789">現在、サポートされているイベントの種類は `Office.EventType.AttachmentsChanged`、`Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged`、`Office.EventType.RecurrenceChanged` です。</span><span class="sxs-lookup"><span data-stu-id="162c4-789">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="162c4-790">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="162c4-790">Parameters:</span></span>

| <span data-ttu-id="162c4-791">名前</span><span class="sxs-lookup"><span data-stu-id="162c4-791">Name</span></span> | <span data-ttu-id="162c4-792">型</span><span class="sxs-lookup"><span data-stu-id="162c4-792">Type</span></span> | <span data-ttu-id="162c4-793">属性</span><span class="sxs-lookup"><span data-stu-id="162c4-793">Attributes</span></span> | <span data-ttu-id="162c4-794">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-794">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="162c4-795">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="162c4-795">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="162c4-796">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="162c4-796">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="162c4-797">Function</span><span class="sxs-lookup"><span data-stu-id="162c4-797">Function</span></span> || <span data-ttu-id="162c4-p138">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="162c4-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="162c4-801">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-801">Object</span></span> | <span data-ttu-id="162c4-802">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-802">&lt;optional&gt;</span></span> | <span data-ttu-id="162c4-803">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="162c4-803">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="162c4-804">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-804">Object</span></span> | <span data-ttu-id="162c4-805">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-805">&lt;optional&gt;</span></span> | <span data-ttu-id="162c4-806">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="162c4-806">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="162c4-807">function</span><span class="sxs-lookup"><span data-stu-id="162c4-807">function</span></span>| <span data-ttu-id="162c4-808">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-808">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-809">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-809">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="162c4-810">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-810">Requirements</span></span>

|<span data-ttu-id="162c4-811">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-811">Requirement</span></span>| <span data-ttu-id="162c4-812">値</span><span class="sxs-lookup"><span data-stu-id="162c4-812">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-813">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-813">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="162c4-814">1.7</span><span class="sxs-lookup"><span data-stu-id="162c4-814">-17</span></span> |
|[<span data-ttu-id="162c4-815">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-815">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="162c4-816">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-816">ReadItem</span></span> |
|[<span data-ttu-id="162c4-817">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-817">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="162c4-818">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="162c4-818">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="162c4-819">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="162c4-819">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="162c4-820">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="162c4-820">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="162c4-p139">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="162c4-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="162c4-824">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="162c4-824">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="162c4-825">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、推奨されていません。</span><span class="sxs-lookup"><span data-stu-id="162c4-825">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="162c4-826">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="162c4-826">Parameters:</span></span>

|<span data-ttu-id="162c4-827">名前</span><span class="sxs-lookup"><span data-stu-id="162c4-827">Name</span></span>|<span data-ttu-id="162c4-828">型</span><span class="sxs-lookup"><span data-stu-id="162c4-828">Type</span></span>|<span data-ttu-id="162c4-829">属性</span><span class="sxs-lookup"><span data-stu-id="162c4-829">Attributes</span></span>|<span data-ttu-id="162c4-830">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-830">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="162c4-831">String</span><span class="sxs-lookup"><span data-stu-id="162c4-831">String</span></span>||<span data-ttu-id="162c4-p140">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="162c4-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="162c4-834">String</span><span class="sxs-lookup"><span data-stu-id="162c4-834">String</span></span>||<span data-ttu-id="162c4-p141">添付するアイテムの件名。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="162c4-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="162c4-837">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-837">Object</span></span>|<span data-ttu-id="162c4-838">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-838">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-839">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="162c4-839">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="162c4-840">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-840">Object</span></span>|<span data-ttu-id="162c4-841">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-841">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-842">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="162c4-842">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="162c4-843">function</span><span class="sxs-lookup"><span data-stu-id="162c4-843">function</span></span>|<span data-ttu-id="162c4-844">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-844">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-845">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-845">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="162c4-846">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-846">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="162c4-847">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="162c4-847">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="162c4-848">エラー</span><span class="sxs-lookup"><span data-stu-id="162c4-848">Errors</span></span>

|<span data-ttu-id="162c4-849">エラー コード</span><span class="sxs-lookup"><span data-stu-id="162c4-849">Error code</span></span>|<span data-ttu-id="162c4-850">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-850">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="162c4-851">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="162c4-851">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="162c4-852">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-852">Requirements</span></span>

|<span data-ttu-id="162c4-853">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-853">Requirement</span></span>|<span data-ttu-id="162c4-854">値</span><span class="sxs-lookup"><span data-stu-id="162c4-854">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-855">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-855">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-856">1.1</span><span class="sxs-lookup"><span data-stu-id="162c4-856">1.1</span></span>|
|[<span data-ttu-id="162c4-857">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-857">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-858">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="162c4-858">ReadWriteItem</span></span>|
|[<span data-ttu-id="162c4-859">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-859">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-860">Compose</span><span class="sxs-lookup"><span data-stu-id="162c4-860">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-861">例</span><span class="sxs-lookup"><span data-stu-id="162c4-861">Example</span></span>

<span data-ttu-id="162c4-862">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-862">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```javascript
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

####  <a name="close"></a><span data-ttu-id="162c4-863">close()</span><span class="sxs-lookup"><span data-stu-id="162c4-863">close()</span></span>

<span data-ttu-id="162c4-864">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="162c4-864">Closes the current item that is being composed.</span></span>

<span data-ttu-id="162c4-p142">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="162c4-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="162c4-867">Outlook on the web では、アイテムが予定であり、`saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、またはキャンセルするようにユーザーは求められます。</span><span class="sxs-lookup"><span data-stu-id="162c4-867">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="162c4-868">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="162c4-868">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-869">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-869">Requirements</span></span>

|<span data-ttu-id="162c4-870">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-870">Requirement</span></span>|<span data-ttu-id="162c4-871">値</span><span class="sxs-lookup"><span data-stu-id="162c4-871">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-872">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-872">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-873">1.3</span><span class="sxs-lookup"><span data-stu-id="162c4-873">1.3</span></span>|
|[<span data-ttu-id="162c4-874">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-874">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-875">制限あり</span><span class="sxs-lookup"><span data-stu-id="162c4-875">Restricted</span></span>|
|[<span data-ttu-id="162c4-876">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-876">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-877">Compose</span><span class="sxs-lookup"><span data-stu-id="162c4-877">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="162c4-878">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="162c4-878">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="162c4-879">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-879">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="162c4-880">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="162c4-880">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="162c4-881">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-881">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="162c4-882">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="162c4-882">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="162c4-p143">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="162c4-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="162c4-886">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="162c4-886">Parameters:</span></span>

|<span data-ttu-id="162c4-887">名前</span><span class="sxs-lookup"><span data-stu-id="162c4-887">Name</span></span>|<span data-ttu-id="162c4-888">型</span><span class="sxs-lookup"><span data-stu-id="162c4-888">Type</span></span>|<span data-ttu-id="162c4-889">属性</span><span class="sxs-lookup"><span data-stu-id="162c4-889">Attributes</span></span>|<span data-ttu-id="162c4-890">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-890">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="162c4-891">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="162c4-891">String &#124; Object</span></span>||<span data-ttu-id="162c4-p144">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="162c4-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="162c4-894">**または**</span><span class="sxs-lookup"><span data-stu-id="162c4-894">**OR**</span></span><br/><span data-ttu-id="162c4-p145">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="162c4-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="162c4-897">String</span><span class="sxs-lookup"><span data-stu-id="162c4-897">String</span></span>|<span data-ttu-id="162c4-898">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-898">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="162c4-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="162c4-901">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-901">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="162c4-902">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-902">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-903">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="162c4-903">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="162c4-904">String</span><span class="sxs-lookup"><span data-stu-id="162c4-904">String</span></span>||<span data-ttu-id="162c4-p147">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="162c4-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="162c4-907">String</span><span class="sxs-lookup"><span data-stu-id="162c4-907">String</span></span>||<span data-ttu-id="162c4-908">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="162c4-908">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="162c4-909">String</span><span class="sxs-lookup"><span data-stu-id="162c4-909">String</span></span>||<span data-ttu-id="162c4-p148">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="162c4-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="162c4-912">ブール値</span><span class="sxs-lookup"><span data-stu-id="162c4-912">Boolean</span></span>||<span data-ttu-id="162c4-p149">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="162c4-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="162c4-915">String</span><span class="sxs-lookup"><span data-stu-id="162c4-915">String</span></span>||<span data-ttu-id="162c4-p150">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="162c4-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="162c4-919">function</span><span class="sxs-lookup"><span data-stu-id="162c4-919">function</span></span>|<span data-ttu-id="162c4-920">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-920">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-921">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-921">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="162c4-922">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-922">Requirements</span></span>

|<span data-ttu-id="162c4-923">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-923">Requirement</span></span>|<span data-ttu-id="162c4-924">値</span><span class="sxs-lookup"><span data-stu-id="162c4-924">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-925">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-925">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-926">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-926">1.0</span></span>|
|[<span data-ttu-id="162c4-927">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-927">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-928">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-928">ReadItem</span></span>|
|[<span data-ttu-id="162c4-929">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-929">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-930">Read</span><span class="sxs-lookup"><span data-stu-id="162c4-930">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="162c4-931">例</span><span class="sxs-lookup"><span data-stu-id="162c4-931">Examples</span></span>

<span data-ttu-id="162c4-932">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="162c4-932">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="162c4-933">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="162c4-933">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="162c4-934">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="162c4-934">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="162c4-935">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="162c4-935">Reply with a body and a file attachment.</span></span>

```javascript
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

<span data-ttu-id="162c4-936">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="162c4-936">Reply with a body and an item attachment.</span></span>

```javascript
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

<span data-ttu-id="162c4-937">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="162c4-937">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```javascript
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="162c4-938">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="162c4-938">displayReplyForm(formData)</span></span>

<span data-ttu-id="162c4-939">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-939">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="162c4-940">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="162c4-940">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="162c4-941">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-941">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="162c4-942">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="162c4-942">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="162c4-p151">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="162c4-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="162c4-946">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="162c4-946">Parameters:</span></span>

|<span data-ttu-id="162c4-947">名前</span><span class="sxs-lookup"><span data-stu-id="162c4-947">Name</span></span>|<span data-ttu-id="162c4-948">型</span><span class="sxs-lookup"><span data-stu-id="162c4-948">Type</span></span>|<span data-ttu-id="162c4-949">属性</span><span class="sxs-lookup"><span data-stu-id="162c4-949">Attributes</span></span>|<span data-ttu-id="162c4-950">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-950">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="162c4-951">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="162c4-951">String &#124; Object</span></span>||<span data-ttu-id="162c4-p152">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="162c4-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="162c4-954">**または**</span><span class="sxs-lookup"><span data-stu-id="162c4-954">**OR**</span></span><br/><span data-ttu-id="162c4-p153">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="162c4-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="162c4-957">String</span><span class="sxs-lookup"><span data-stu-id="162c4-957">String</span></span>|<span data-ttu-id="162c4-958">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-958">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-p154">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="162c4-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="162c4-961">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-961">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="162c4-962">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-962">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-963">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="162c4-963">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="162c4-964">String</span><span class="sxs-lookup"><span data-stu-id="162c4-964">String</span></span>||<span data-ttu-id="162c4-p155">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="162c4-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="162c4-967">String</span><span class="sxs-lookup"><span data-stu-id="162c4-967">String</span></span>||<span data-ttu-id="162c4-968">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="162c4-968">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="162c4-969">String</span><span class="sxs-lookup"><span data-stu-id="162c4-969">String</span></span>||<span data-ttu-id="162c4-p156">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="162c4-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="162c4-972">ブール値</span><span class="sxs-lookup"><span data-stu-id="162c4-972">Boolean</span></span>||<span data-ttu-id="162c4-p157">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="162c4-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="162c4-975">String</span><span class="sxs-lookup"><span data-stu-id="162c4-975">String</span></span>||<span data-ttu-id="162c4-p158">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="162c4-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="162c4-979">function</span><span class="sxs-lookup"><span data-stu-id="162c4-979">function</span></span>|<span data-ttu-id="162c4-980">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-980">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-981">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-981">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="162c4-982">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-982">Requirements</span></span>

|<span data-ttu-id="162c4-983">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-983">Requirement</span></span>|<span data-ttu-id="162c4-984">値</span><span class="sxs-lookup"><span data-stu-id="162c4-984">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-985">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-985">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-986">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-986">1.0</span></span>|
|[<span data-ttu-id="162c4-987">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-987">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-988">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-988">ReadItem</span></span>|
|[<span data-ttu-id="162c4-989">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-989">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-990">Read</span><span class="sxs-lookup"><span data-stu-id="162c4-990">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="162c4-991">例</span><span class="sxs-lookup"><span data-stu-id="162c4-991">Examples</span></span>

<span data-ttu-id="162c4-992">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="162c4-992">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="162c4-993">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="162c4-993">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="162c4-994">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="162c4-994">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="162c4-995">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="162c4-995">Reply with a body and a file attachment.</span></span>

```javascript
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

<span data-ttu-id="162c4-996">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="162c4-996">Reply with a body and an item attachment.</span></span>

```javascript
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

<span data-ttu-id="162c4-997">本文、ファイルの添付ファイル、アイテムの添付ファイル、コールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="162c4-997">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```javascript
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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="162c4-998">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="162c4-998">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="162c4-999">メッセージまたは予定から指定の添付ファイルを取得し、それを `AttachmentContent` オブジェクトとして返されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-999">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="162c4-1000">`getAttachmentContentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから取得します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1000">The `getAttachmentContentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="162c4-1001">ベスト プラクティスとして、識別子を使用し、`getAttachmentsAsync` または `item.attachments` 呼び出しで attachmentIds を取得した同じセッションで添付ファイルを取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="162c4-1001">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="162c4-1002">Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="162c4-1002">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="162c4-1003">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始し、フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1003">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="162c4-1004">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="162c4-1004">Parameters:</span></span>

|<span data-ttu-id="162c4-1005">名前</span><span class="sxs-lookup"><span data-stu-id="162c4-1005">Name</span></span>|<span data-ttu-id="162c4-1006">型</span><span class="sxs-lookup"><span data-stu-id="162c4-1006">Type</span></span>|<span data-ttu-id="162c4-1007">属性</span><span class="sxs-lookup"><span data-stu-id="162c4-1007">Attributes</span></span>|<span data-ttu-id="162c4-1008">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-1008">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="162c4-1009">String</span><span class="sxs-lookup"><span data-stu-id="162c4-1009">String</span></span>||<span data-ttu-id="162c4-1010">取得する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="162c4-1010">The identifier of the attachment you want to get.</span></span> <span data-ttu-id="162c4-1011">文字列の最大の長さは 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="162c4-1011">The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="162c4-1012">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-1012">Object</span></span>|<span data-ttu-id="162c4-1013">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1013">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1014">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="162c4-1014">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="162c4-1015">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-1015">Object</span></span>|<span data-ttu-id="162c4-1016">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1016">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1017">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1017">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="162c4-1018">function</span><span class="sxs-lookup"><span data-stu-id="162c4-1018">function</span></span>|<span data-ttu-id="162c4-1019">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1019">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1020">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1020">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="162c4-1021">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1021">Requirements</span></span>

|<span data-ttu-id="162c4-1022">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1022">Requirement</span></span>|<span data-ttu-id="162c4-1023">値</span><span class="sxs-lookup"><span data-stu-id="162c4-1023">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-1024">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-1024">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-1025">プレビュー</span><span class="sxs-lookup"><span data-stu-id="162c4-1025">Preview</span></span>|
|[<span data-ttu-id="162c4-1026">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-1026">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-1027">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-1027">ReadItem</span></span>|
|[<span data-ttu-id="162c4-1028">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-1028">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-1029">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="162c4-1029">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="162c4-1030">戻り値:</span><span class="sxs-lookup"><span data-stu-id="162c4-1030">Returns:</span></span>

<span data-ttu-id="162c4-1031">型: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="162c4-1031">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="162c4-1032">例</span><span class="sxs-lookup"><span data-stu-id="162c4-1032">Example</span></span>

```javascript
var item = Office.context.mailbox.item;
var listOfAttachments = [];
item.getAttachmentsAsync(callback);
function callback(result) {
    if (result.value.length > 0) {
        for (i = 0 ; i < result.value.length ; i++) {
            var options = {asyncContext: {type: result.value[i].attachmentType}};
            getAttachmentContentAsync(result.value[i].id, options, handleAttachmentsCallback);  
        }
    }
}

function handleAttachmentsCallback(result) {
    // parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file
    if (result.format == Office.MailboxEnums.AttachmentContentFormat.Base64) {
        // handle file attachment
    }
    else if (result.format == Office.MailboxEnums.AttachmentContentFormat.Eml) {
        // handle item attachment
    }
    else if (result.format == Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
        // handle .icalender attachment
    }
    else {
        // handle cloud attachment  
    }
}
```

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="162c4-1033">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="162c4-1033">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="162c4-1034">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1034">Gets the item's attachments as an array.</span></span> <span data-ttu-id="162c4-1035">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="162c4-1035">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="162c4-1036">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="162c4-1036">Parameters:</span></span>

|<span data-ttu-id="162c4-1037">名前</span><span class="sxs-lookup"><span data-stu-id="162c4-1037">Name</span></span>|<span data-ttu-id="162c4-1038">型</span><span class="sxs-lookup"><span data-stu-id="162c4-1038">Type</span></span>|<span data-ttu-id="162c4-1039">属性</span><span class="sxs-lookup"><span data-stu-id="162c4-1039">Attributes</span></span>|<span data-ttu-id="162c4-1040">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-1040">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="162c4-1041">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-1041">Object</span></span>|<span data-ttu-id="162c4-1042">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1043">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="162c4-1043">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="162c4-1044">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-1044">Object</span></span>|<span data-ttu-id="162c4-1045">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1046">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1046">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="162c4-1047">function</span><span class="sxs-lookup"><span data-stu-id="162c4-1047">function</span></span>|<span data-ttu-id="162c4-1048">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1048">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1049">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1049">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="162c4-1050">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1050">Requirements</span></span>

|<span data-ttu-id="162c4-1051">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1051">Requirement</span></span>|<span data-ttu-id="162c4-1052">値</span><span class="sxs-lookup"><span data-stu-id="162c4-1052">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-1053">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-1053">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-1054">プレビュー</span><span class="sxs-lookup"><span data-stu-id="162c4-1054">Preview</span></span>|
|[<span data-ttu-id="162c4-1055">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-1055">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-1056">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-1056">ReadItem</span></span>|
|[<span data-ttu-id="162c4-1057">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-1057">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-1058">Compose</span><span class="sxs-lookup"><span data-stu-id="162c4-1058">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="162c4-1059">戻り値:</span><span class="sxs-lookup"><span data-stu-id="162c4-1059">Returns:</span></span>

<span data-ttu-id="162c4-1060">型: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="162c4-1060">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="162c4-1061">例</span><span class="sxs-lookup"><span data-stu-id="162c4-1061">Example</span></span>

<span data-ttu-id="162c4-1062">次の例では、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1062">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);  
function callback(result) {
    if (result.value.length > 0) {
        for (i = 0 ; i < result.value.length ; i++) {
            var _att = result.value [i];
            outputString += "<BR>" + i + ". Name: ";
            outputString += _att.name;
            outputString += "<BR>ID: " + _att.id;
            outputString += "<BR>contentType: " + _att.contentType;
            outputString += "<BR>size: " + _att.size;
            outputString += "<BR>attachmentType: " + _att.attachmentType;
            outputString += "<BR>isInline: " + _att.isInline;
        }
    }
}
```

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="162c4-1063">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="162c4-1063">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="162c4-1064">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1064">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="162c4-1065">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="162c4-1065">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-1066">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1066">Requirements</span></span>

|<span data-ttu-id="162c4-1067">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1067">Requirement</span></span>|<span data-ttu-id="162c4-1068">値</span><span class="sxs-lookup"><span data-stu-id="162c4-1068">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-1069">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-1069">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-1070">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-1070">1.0</span></span>|
|[<span data-ttu-id="162c4-1071">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-1071">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-1072">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-1072">ReadItem</span></span>|
|[<span data-ttu-id="162c4-1073">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-1073">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-1074">Read</span><span class="sxs-lookup"><span data-stu-id="162c4-1074">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="162c4-1075">戻り値:</span><span class="sxs-lookup"><span data-stu-id="162c4-1075">Returns:</span></span>

<span data-ttu-id="162c4-1076">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="162c4-1076">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="162c4-1077">例</span><span class="sxs-lookup"><span data-stu-id="162c4-1077">Example</span></span>

<span data-ttu-id="162c4-1078">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="162c4-1078">The following example accesses the contacts entities on the current item.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="162c4-1079">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="162c4-1079">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="162c4-1080">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1080">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="162c4-1081">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="162c4-1081">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="162c4-1082">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="162c4-1082">Parameters:</span></span>

|<span data-ttu-id="162c4-1083">名前</span><span class="sxs-lookup"><span data-stu-id="162c4-1083">Name</span></span>|<span data-ttu-id="162c4-1084">型</span><span class="sxs-lookup"><span data-stu-id="162c4-1084">Type</span></span>|<span data-ttu-id="162c4-1085">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-1085">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="162c4-1086">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="162c4-1086">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="162c4-1087">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="162c4-1087">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="162c4-1088">Requirements</span><span class="sxs-lookup"><span data-stu-id="162c4-1088">Requirements</span></span>

|<span data-ttu-id="162c4-1089">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1089">Requirement</span></span>|<span data-ttu-id="162c4-1090">値</span><span class="sxs-lookup"><span data-stu-id="162c4-1090">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-1091">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-1091">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-1092">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-1092">1.0</span></span>|
|[<span data-ttu-id="162c4-1093">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-1093">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-1094">制限あり</span><span class="sxs-lookup"><span data-stu-id="162c4-1094">Restricted</span></span>|
|[<span data-ttu-id="162c4-1095">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-1095">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-1096">Read</span><span class="sxs-lookup"><span data-stu-id="162c4-1096">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="162c4-1097">戻り値:</span><span class="sxs-lookup"><span data-stu-id="162c4-1097">Returns:</span></span>

<span data-ttu-id="162c4-1098">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1098">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="162c4-1099">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1099">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="162c4-1100">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="162c4-1100">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="162c4-1101">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="162c4-1101">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="162c4-1102">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="162c4-1102">Value of `entityType`</span></span>|<span data-ttu-id="162c4-1103">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="162c4-1103">Type of objects in returned array</span></span>|<span data-ttu-id="162c4-1104">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="162c4-1104">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="162c4-1105">文字列</span><span class="sxs-lookup"><span data-stu-id="162c4-1105">String</span></span>|<span data-ttu-id="162c4-1106">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="162c4-1106">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="162c4-1107">連絡先</span><span class="sxs-lookup"><span data-stu-id="162c4-1107">Contact</span></span>|<span data-ttu-id="162c4-1108">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="162c4-1108">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="162c4-1109">文字列</span><span class="sxs-lookup"><span data-stu-id="162c4-1109">String</span></span>|<span data-ttu-id="162c4-1110">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="162c4-1110">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="162c4-1111">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="162c4-1111">MeetingSuggestion</span></span>|<span data-ttu-id="162c4-1112">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="162c4-1112">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="162c4-1113">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="162c4-1113">PhoneNumber</span></span>|<span data-ttu-id="162c4-1114">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="162c4-1114">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="162c4-1115">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="162c4-1115">TaskSuggestion</span></span>|<span data-ttu-id="162c4-1116">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="162c4-1116">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="162c4-1117">文字列</span><span class="sxs-lookup"><span data-stu-id="162c4-1117">String</span></span>|<span data-ttu-id="162c4-1118">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="162c4-1118">**Restricted**</span></span>|

<span data-ttu-id="162c4-1119">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="162c4-1119">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="162c4-1120">例</span><span class="sxs-lookup"><span data-stu-id="162c4-1120">Example</span></span>

<span data-ttu-id="162c4-1121">次の例は、現在のアイテムの本文にある郵便アドレスを表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1121">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

```javascript
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="162c4-1122">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="162c4-1122">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="162c4-1123">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1123">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="162c4-1124">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="162c4-1124">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="162c4-1125">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1125">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="162c4-1126">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="162c4-1126">Parameters:</span></span>

|<span data-ttu-id="162c4-1127">名前</span><span class="sxs-lookup"><span data-stu-id="162c4-1127">Name</span></span>|<span data-ttu-id="162c4-1128">型</span><span class="sxs-lookup"><span data-stu-id="162c4-1128">Type</span></span>|<span data-ttu-id="162c4-1129">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-1129">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="162c4-1130">String</span><span class="sxs-lookup"><span data-stu-id="162c4-1130">String</span></span>|<span data-ttu-id="162c4-1131">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="162c4-1131">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="162c4-1132">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1132">Requirements</span></span>

|<span data-ttu-id="162c4-1133">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1133">Requirement</span></span>|<span data-ttu-id="162c4-1134">値</span><span class="sxs-lookup"><span data-stu-id="162c4-1134">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-1135">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-1135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-1136">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-1136">1.0</span></span>|
|[<span data-ttu-id="162c4-1137">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-1137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-1138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-1138">ReadItem</span></span>|
|[<span data-ttu-id="162c4-1139">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-1139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-1140">Read</span><span class="sxs-lookup"><span data-stu-id="162c4-1140">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="162c4-1141">戻り値:</span><span class="sxs-lookup"><span data-stu-id="162c4-1141">Returns:</span></span>

<span data-ttu-id="162c4-p163">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-p163">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="162c4-1144">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="162c4-1144">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="162c4-1145">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="162c4-1145">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="162c4-1146">アドインが[操作可能メッセージによってアクティブ化](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを取得します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1146">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="162c4-1147">このメソッドは、Outlook 2016 for Windows 以降 (16.0.8413.1000 以降のクイック実行バージョン) および Outlook on the web for Office 365 でのみサポートされます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1147">Note: This method is only supported by Outlook 2016 for Windows (Click-to-Run versions greater than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="162c4-1148">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="162c4-1148">Parameters:</span></span>
|<span data-ttu-id="162c4-1149">名前</span><span class="sxs-lookup"><span data-stu-id="162c4-1149">Name</span></span>|<span data-ttu-id="162c4-1150">型</span><span class="sxs-lookup"><span data-stu-id="162c4-1150">Type</span></span>|<span data-ttu-id="162c4-1151">属性</span><span class="sxs-lookup"><span data-stu-id="162c4-1151">Attributes</span></span>|<span data-ttu-id="162c4-1152">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-1152">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="162c4-1153">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-1153">Object</span></span>|<span data-ttu-id="162c4-1154">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1154">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1155">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="162c4-1155">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="162c4-1156">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-1156">Object</span></span>|<span data-ttu-id="162c4-1157">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1158">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1158">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="162c4-1159">function</span><span class="sxs-lookup"><span data-stu-id="162c4-1159">function</span></span>|<span data-ttu-id="162c4-1160">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1161">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1161">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="162c4-1162">成功すると、初期化データが文字列として `asyncResult.value` プロパティで指定されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1162">On success, the intialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="162c4-1163">初期化コンテキストがない場合、`asyncResult` オブジェクトには、`code` プロパティが `9020`、`name` プロパティが `GenericResponseError` に設定された `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1163">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="162c4-1164">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1164">Requirements</span></span>

|<span data-ttu-id="162c4-1165">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1165">Requirement</span></span>|<span data-ttu-id="162c4-1166">値</span><span class="sxs-lookup"><span data-stu-id="162c4-1166">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-1167">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-1167">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-1168">プレビュー</span><span class="sxs-lookup"><span data-stu-id="162c4-1168">Preview</span></span>|
|[<span data-ttu-id="162c4-1169">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-1169">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-1170">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-1170">ReadItem</span></span>|
|[<span data-ttu-id="162c4-1171">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-1171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-1172">Read</span><span class="sxs-lookup"><span data-stu-id="162c4-1172">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-1173">例</span><span class="sxs-lookup"><span data-stu-id="162c4-1173">Example</span></span>

```javascript
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

#### <a name="getregexmatches--object"></a><span data-ttu-id="162c4-1174">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="162c4-1174">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="162c4-1175">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1175">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="162c4-1176">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="162c4-1176">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="162c4-p164">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="162c4-p164">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="162c4-1180">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="162c4-1180">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="162c4-1181">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="162c4-1181">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="162c4-p165">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="162c4-p165">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-1185">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1185">Requirements</span></span>

|<span data-ttu-id="162c4-1186">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1186">Requirement</span></span>|<span data-ttu-id="162c4-1187">値</span><span class="sxs-lookup"><span data-stu-id="162c4-1187">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-1188">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-1188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-1189">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-1189">1.0</span></span>|
|[<span data-ttu-id="162c4-1190">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-1190">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-1191">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-1191">ReadItem</span></span>|
|[<span data-ttu-id="162c4-1192">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-1192">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-1193">Read</span><span class="sxs-lookup"><span data-stu-id="162c4-1193">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="162c4-1194">戻り値:</span><span class="sxs-lookup"><span data-stu-id="162c4-1194">Returns:</span></span>

<span data-ttu-id="162c4-p166">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="162c4-p166">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="162c4-1197">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="162c4-1197">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="162c4-1198">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-1198">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="162c4-1199">例</span><span class="sxs-lookup"><span data-stu-id="162c4-1199">Example</span></span>

<span data-ttu-id="162c4-1200">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="162c4-1200">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="162c4-1201">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="162c4-1201">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="162c4-1202">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定の正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1202">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="162c4-1203">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="162c4-1203">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="162c4-1204">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="162c4-1204">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="162c4-p167">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="162c4-p167">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="162c4-1207">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="162c4-1207">Parameters:</span></span>

|<span data-ttu-id="162c4-1208">名前</span><span class="sxs-lookup"><span data-stu-id="162c4-1208">Name</span></span>|<span data-ttu-id="162c4-1209">型</span><span class="sxs-lookup"><span data-stu-id="162c4-1209">Type</span></span>|<span data-ttu-id="162c4-1210">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-1210">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="162c4-1211">String</span><span class="sxs-lookup"><span data-stu-id="162c4-1211">String</span></span>|<span data-ttu-id="162c4-1212">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="162c4-1212">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="162c4-1213">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1213">Requirements</span></span>

|<span data-ttu-id="162c4-1214">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1214">Requirement</span></span>|<span data-ttu-id="162c4-1215">値</span><span class="sxs-lookup"><span data-stu-id="162c4-1215">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-1216">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-1216">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-1217">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-1217">1.0</span></span>|
|[<span data-ttu-id="162c4-1218">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-1218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-1219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-1219">ReadItem</span></span>|
|[<span data-ttu-id="162c4-1220">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-1220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-1221">Read</span><span class="sxs-lookup"><span data-stu-id="162c4-1221">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="162c4-1222">戻り値:</span><span class="sxs-lookup"><span data-stu-id="162c4-1222">Returns:</span></span>

<span data-ttu-id="162c4-1223">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="162c4-1223">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="162c4-1224">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="162c4-1224">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="162c4-1225">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="162c4-1225">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="162c4-1226">例</span><span class="sxs-lookup"><span data-stu-id="162c4-1226">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="162c4-1227">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="162c4-1227">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="162c4-1228">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1228">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="162c4-p168">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-p168">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="162c4-1231">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="162c4-1231">Parameters:</span></span>

|<span data-ttu-id="162c4-1232">名前</span><span class="sxs-lookup"><span data-stu-id="162c4-1232">Name</span></span>|<span data-ttu-id="162c4-1233">型</span><span class="sxs-lookup"><span data-stu-id="162c4-1233">Type</span></span>|<span data-ttu-id="162c4-1234">属性</span><span class="sxs-lookup"><span data-stu-id="162c4-1234">Attributes</span></span>|<span data-ttu-id="162c4-1235">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-1235">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="162c4-1236">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="162c4-1236">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="162c4-p169">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="162c4-p169">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="162c4-1240">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-1240">Object</span></span>|<span data-ttu-id="162c4-1241">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1241">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1242">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="162c4-1242">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="162c4-1243">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-1243">Object</span></span>|<span data-ttu-id="162c4-1244">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1244">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1245">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1245">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="162c4-1246">function</span><span class="sxs-lookup"><span data-stu-id="162c4-1246">function</span></span>||<span data-ttu-id="162c4-1247">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1247">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="162c4-1248">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1248">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="162c4-1249">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="162c4-1249">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="162c4-1250">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1250">Requirements</span></span>

|<span data-ttu-id="162c4-1251">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1251">Requirement</span></span>|<span data-ttu-id="162c4-1252">値</span><span class="sxs-lookup"><span data-stu-id="162c4-1252">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-1253">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-1253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-1254">1.2</span><span class="sxs-lookup"><span data-stu-id="162c4-1254">1.2</span></span>|
|[<span data-ttu-id="162c4-1255">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-1255">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-1256">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="162c4-1256">ReadWriteItem</span></span>|
|[<span data-ttu-id="162c4-1257">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-1257">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-1258">Compose</span><span class="sxs-lookup"><span data-stu-id="162c4-1258">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="162c4-1259">戻り値:</span><span class="sxs-lookup"><span data-stu-id="162c4-1259">Returns:</span></span>

<span data-ttu-id="162c4-1260">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="162c4-1260">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="162c4-1261">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="162c4-1261">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="162c4-1262">String</span><span class="sxs-lookup"><span data-stu-id="162c4-1262">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="162c4-1263">例</span><span class="sxs-lookup"><span data-stu-id="162c4-1263">Example</span></span>

```javascript
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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="162c4-1264">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="162c4-1264">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="162c4-p171">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-p171">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="162c4-1267">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="162c4-1267">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-1268">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1268">Requirements</span></span>

|<span data-ttu-id="162c4-1269">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1269">Requirement</span></span>|<span data-ttu-id="162c4-1270">値</span><span class="sxs-lookup"><span data-stu-id="162c4-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-1271">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-1272">1.6</span><span class="sxs-lookup"><span data-stu-id="162c4-1272">-16</span></span>|
|[<span data-ttu-id="162c4-1273">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-1273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-1274">ReadItem</span></span>|
|[<span data-ttu-id="162c4-1275">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-1275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-1276">Read</span><span class="sxs-lookup"><span data-stu-id="162c4-1276">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="162c4-1277">戻り値:</span><span class="sxs-lookup"><span data-stu-id="162c4-1277">Returns:</span></span>

<span data-ttu-id="162c4-1278">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="162c4-1278">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="162c4-1279">例</span><span class="sxs-lookup"><span data-stu-id="162c4-1279">Example</span></span>

<span data-ttu-id="162c4-1280">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="162c4-1280">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="162c4-1281">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="162c4-1281">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="162c4-p172">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-p172">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="162c4-1284">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="162c4-1284">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="162c4-p173">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="162c4-p173">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="162c4-1288">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="162c4-1288">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="162c4-1289">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="162c4-1289">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="162c4-p174">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="162c4-p174">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="162c4-1293">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1293">Requirements</span></span>

|<span data-ttu-id="162c4-1294">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1294">Requirement</span></span>|<span data-ttu-id="162c4-1295">値</span><span class="sxs-lookup"><span data-stu-id="162c4-1295">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-1296">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-1296">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-1297">1.6</span><span class="sxs-lookup"><span data-stu-id="162c4-1297">-16</span></span>|
|[<span data-ttu-id="162c4-1298">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-1298">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-1299">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-1299">ReadItem</span></span>|
|[<span data-ttu-id="162c4-1300">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-1300">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-1301">Read</span><span class="sxs-lookup"><span data-stu-id="162c4-1301">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="162c4-1302">戻り値:</span><span class="sxs-lookup"><span data-stu-id="162c4-1302">Returns:</span></span>

<span data-ttu-id="162c4-p175">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="162c4-p175">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="162c4-1305">例</span><span class="sxs-lookup"><span data-stu-id="162c4-1305">Example</span></span>

<span data-ttu-id="162c4-1306">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="162c4-1306">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="162c4-1307">getSharedPropertiesAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="162c4-1307">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="162c4-1308">共有フォルダー、カレンダー、メールボックスで選択した予定またはメッセージのプロパティを取得します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1308">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="162c4-1309">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="162c4-1309">Parameters:</span></span>

|<span data-ttu-id="162c4-1310">名前</span><span class="sxs-lookup"><span data-stu-id="162c4-1310">Name</span></span>|<span data-ttu-id="162c4-1311">型</span><span class="sxs-lookup"><span data-stu-id="162c4-1311">Type</span></span>|<span data-ttu-id="162c4-1312">属性</span><span class="sxs-lookup"><span data-stu-id="162c4-1312">Attributes</span></span>|<span data-ttu-id="162c4-1313">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-1313">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="162c4-1314">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-1314">Object</span></span>|<span data-ttu-id="162c4-1315">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1315">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1316">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="162c4-1316">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="162c4-1317">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-1317">Object</span></span>|<span data-ttu-id="162c4-1318">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1318">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1319">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1319">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="162c4-1320">function</span><span class="sxs-lookup"><span data-stu-id="162c4-1320">function</span></span>||<span data-ttu-id="162c4-1321">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1321">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="162c4-1322">共有プロパティは `asyncResult.value` プロパティの [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1322">The custom properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="162c4-1323">このオブジェクトは、アイテムの共有プロパティの取得に使用できます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1323">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="162c4-1324">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1324">Requirements</span></span>

|<span data-ttu-id="162c4-1325">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1325">Requirement</span></span>|<span data-ttu-id="162c4-1326">値</span><span class="sxs-lookup"><span data-stu-id="162c4-1326">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-1327">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-1327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-1328">プレビュー</span><span class="sxs-lookup"><span data-stu-id="162c4-1328">Preview</span></span>|
|[<span data-ttu-id="162c4-1329">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-1329">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-1330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-1330">ReadItem</span></span>|
|[<span data-ttu-id="162c4-1331">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-1331">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-1332">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="162c4-1332">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-1333">例</span><span class="sxs-lookup"><span data-stu-id="162c4-1333">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="162c4-1334">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="162c4-1334">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="162c4-1335">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1335">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="162c4-p177">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="162c4-p177">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="162c4-1339">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="162c4-1339">Parameters:</span></span>

|<span data-ttu-id="162c4-1340">名前</span><span class="sxs-lookup"><span data-stu-id="162c4-1340">Name</span></span>|<span data-ttu-id="162c4-1341">型</span><span class="sxs-lookup"><span data-stu-id="162c4-1341">Type</span></span>|<span data-ttu-id="162c4-1342">属性</span><span class="sxs-lookup"><span data-stu-id="162c4-1342">Attributes</span></span>|<span data-ttu-id="162c4-1343">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-1343">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="162c4-1344">function</span><span class="sxs-lookup"><span data-stu-id="162c4-1344">function</span></span>||<span data-ttu-id="162c4-1345">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1345">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="162c4-1346">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1346">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="162c4-1347">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、および削除し、カスタム プロパティに対する変更をサーバーに設定し直すために使用できます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1347">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="162c4-1348">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="162c4-1348">Object</span></span>|<span data-ttu-id="162c4-1349">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1349">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1350">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1350">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="162c4-1351">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1351">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="162c4-1352">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1352">Requirements</span></span>

|<span data-ttu-id="162c4-1353">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1353">Requirement</span></span>|<span data-ttu-id="162c4-1354">値</span><span class="sxs-lookup"><span data-stu-id="162c4-1354">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-1355">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-1355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-1356">1.0</span><span class="sxs-lookup"><span data-stu-id="162c4-1356">1.0</span></span>|
|[<span data-ttu-id="162c4-1357">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-1357">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-1358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-1358">ReadItem</span></span>|
|[<span data-ttu-id="162c4-1359">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-1359">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-1360">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="162c4-1360">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-1361">例</span><span class="sxs-lookup"><span data-stu-id="162c4-1361">Example</span></span>

<span data-ttu-id="162c4-p180">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="162c4-p180">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="162c4-1365">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="162c4-1365">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="162c4-1366">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1366">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="162c4-1367">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1367">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="162c4-1368">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="162c4-1368">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="162c4-1369">Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="162c4-1369">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="162c4-1370">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始し、フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1370">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="162c4-1371">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="162c4-1371">Parameters:</span></span>

|<span data-ttu-id="162c4-1372">名前</span><span class="sxs-lookup"><span data-stu-id="162c4-1372">Name</span></span>|<span data-ttu-id="162c4-1373">型</span><span class="sxs-lookup"><span data-stu-id="162c4-1373">Type</span></span>|<span data-ttu-id="162c4-1374">属性</span><span class="sxs-lookup"><span data-stu-id="162c4-1374">Attributes</span></span>|<span data-ttu-id="162c4-1375">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-1375">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="162c4-1376">String</span><span class="sxs-lookup"><span data-stu-id="162c4-1376">String</span></span>||<span data-ttu-id="162c4-p182">削除する添付ファイルの識別子。文字列の最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="162c4-p182">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="162c4-1379">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-1379">Object</span></span>|<span data-ttu-id="162c4-1380">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1380">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1381">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="162c4-1381">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="162c4-1382">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-1382">Object</span></span>|<span data-ttu-id="162c4-1383">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1383">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1384">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1384">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="162c4-1385">function</span><span class="sxs-lookup"><span data-stu-id="162c4-1385">function</span></span>|<span data-ttu-id="162c4-1386">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1386">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1387">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1387">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="162c4-1388">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1388">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="162c4-1389">エラー</span><span class="sxs-lookup"><span data-stu-id="162c4-1389">Errors</span></span>

|<span data-ttu-id="162c4-1390">エラー コード</span><span class="sxs-lookup"><span data-stu-id="162c4-1390">Error code</span></span>|<span data-ttu-id="162c4-1391">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-1391">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="162c4-1392">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="162c4-1392">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="162c4-1393">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1393">Requirements</span></span>

|<span data-ttu-id="162c4-1394">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1394">Requirement</span></span>|<span data-ttu-id="162c4-1395">値</span><span class="sxs-lookup"><span data-stu-id="162c4-1395">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-1396">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-1396">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-1397">1.1</span><span class="sxs-lookup"><span data-stu-id="162c4-1397">1.1</span></span>|
|[<span data-ttu-id="162c4-1398">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-1398">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-1399">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="162c4-1399">ReadWriteItem</span></span>|
|[<span data-ttu-id="162c4-1400">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-1400">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-1401">Compose</span><span class="sxs-lookup"><span data-stu-id="162c4-1401">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-1402">例</span><span class="sxs-lookup"><span data-stu-id="162c4-1402">Example</span></span>

<span data-ttu-id="162c4-1403">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1403">The following code removes an attachment with an identifier of '0'.</span></span>

```javascript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="162c4-1404">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="162c4-1404">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="162c4-1405">サポートされているイベントのイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1405">Removes an event handler for a</span></span>

<span data-ttu-id="162c4-1406">現在、サポートされているイベントの種類は `Office.EventType.AttachmentsChanged`、`Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged`、`Office.EventType.RecurrenceChanged` です。</span><span class="sxs-lookup"><span data-stu-id="162c4-1406">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="162c4-1407">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="162c4-1407">Parameters:</span></span>

| <span data-ttu-id="162c4-1408">名前</span><span class="sxs-lookup"><span data-stu-id="162c4-1408">Name</span></span> | <span data-ttu-id="162c4-1409">型</span><span class="sxs-lookup"><span data-stu-id="162c4-1409">Type</span></span> | <span data-ttu-id="162c4-1410">属性</span><span class="sxs-lookup"><span data-stu-id="162c4-1410">Attributes</span></span> | <span data-ttu-id="162c4-1411">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-1411">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="162c4-1412">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="162c4-1412">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="162c4-1413">ハンドラーを無効にするイベント。</span><span class="sxs-lookup"><span data-stu-id="162c4-1413">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="162c4-1414">関数</span><span class="sxs-lookup"><span data-stu-id="162c4-1414">Function</span></span> || <span data-ttu-id="162c4-p183">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`removeHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="162c4-p183">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="162c4-1418">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-1418">Object</span></span> | <span data-ttu-id="162c4-1419">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1419">&lt;optional&gt;</span></span> | <span data-ttu-id="162c4-1420">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="162c4-1420">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="162c4-1421">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-1421">Object</span></span> | <span data-ttu-id="162c4-1422">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1422">&lt;optional&gt;</span></span> | <span data-ttu-id="162c4-1423">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1423">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="162c4-1424">function</span><span class="sxs-lookup"><span data-stu-id="162c4-1424">function</span></span>| <span data-ttu-id="162c4-1425">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1425">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1426">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1426">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="162c4-1427">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1427">Requirements</span></span>

|<span data-ttu-id="162c4-1428">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1428">Requirement</span></span>| <span data-ttu-id="162c4-1429">値</span><span class="sxs-lookup"><span data-stu-id="162c4-1429">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-1430">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-1430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="162c4-1431">1.7</span><span class="sxs-lookup"><span data-stu-id="162c4-1431">-17</span></span> |
|[<span data-ttu-id="162c4-1432">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-1432">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="162c4-1433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="162c4-1433">ReadItem</span></span> |
|[<span data-ttu-id="162c4-1434">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-1434">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="162c4-1435">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="162c4-1435">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="162c4-1436">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="162c4-1436">saveAsync([options], callback)</span></span>

<span data-ttu-id="162c4-1437">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1437">Asynchronously saves an item.</span></span>

<span data-ttu-id="162c4-p184">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-p184">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="162c4-1441">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="162c4-1441">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="162c4-1442">アイテムが同期されるまでに、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1442">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="162c4-p186">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-p186">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="162c4-1446">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="162c4-1446">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="162c4-1447">Mac Outlook では、新規作成モードの会議で `saveAsync` をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="162c4-1447">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="162c4-1448">Mac Outlook では、会議で `saveAsync` を呼び出すとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1448">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="162c4-1449">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状またはアップデートが常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1449">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="162c4-1450">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="162c4-1450">Parameters:</span></span>

|<span data-ttu-id="162c4-1451">名前</span><span class="sxs-lookup"><span data-stu-id="162c4-1451">Name</span></span>|<span data-ttu-id="162c4-1452">型</span><span class="sxs-lookup"><span data-stu-id="162c4-1452">Type</span></span>|<span data-ttu-id="162c4-1453">属性</span><span class="sxs-lookup"><span data-stu-id="162c4-1453">Attributes</span></span>|<span data-ttu-id="162c4-1454">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-1454">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="162c4-1455">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-1455">Object</span></span>|<span data-ttu-id="162c4-1456">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1456">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1457">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="162c4-1457">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="162c4-1458">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-1458">Object</span></span>|<span data-ttu-id="162c4-1459">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1459">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1460">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1460">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="162c4-1461">function</span><span class="sxs-lookup"><span data-stu-id="162c4-1461">function</span></span>||<span data-ttu-id="162c4-1462">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1462">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="162c4-1463">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1463">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="162c4-1464">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1464">Requirements</span></span>

|<span data-ttu-id="162c4-1465">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1465">Requirement</span></span>|<span data-ttu-id="162c4-1466">値</span><span class="sxs-lookup"><span data-stu-id="162c4-1466">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-1467">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-1467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-1468">1.3</span><span class="sxs-lookup"><span data-stu-id="162c4-1468">1.3</span></span>|
|[<span data-ttu-id="162c4-1469">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-1469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-1470">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="162c4-1470">ReadWriteItem</span></span>|
|[<span data-ttu-id="162c4-1471">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-1471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-1472">Compose</span><span class="sxs-lookup"><span data-stu-id="162c4-1472">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="162c4-1473">例</span><span class="sxs-lookup"><span data-stu-id="162c4-1473">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="162c4-p188">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="162c4-p188">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="162c4-1476">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="162c4-1476">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="162c4-1477">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="162c4-1477">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="162c4-p189">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="162c4-p189">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="162c4-1481">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="162c4-1481">Parameters:</span></span>

|<span data-ttu-id="162c4-1482">名前</span><span class="sxs-lookup"><span data-stu-id="162c4-1482">Name</span></span>|<span data-ttu-id="162c4-1483">型</span><span class="sxs-lookup"><span data-stu-id="162c4-1483">Type</span></span>|<span data-ttu-id="162c4-1484">属性</span><span class="sxs-lookup"><span data-stu-id="162c4-1484">Attributes</span></span>|<span data-ttu-id="162c4-1485">説明</span><span class="sxs-lookup"><span data-stu-id="162c4-1485">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="162c4-1486">String</span><span class="sxs-lookup"><span data-stu-id="162c4-1486">String</span></span>||<span data-ttu-id="162c4-p190">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="162c4-p190">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="162c4-1490">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-1490">Object</span></span>|<span data-ttu-id="162c4-1491">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1491">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1492">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="162c4-1492">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="162c4-1493">Object</span><span class="sxs-lookup"><span data-stu-id="162c4-1493">Object</span></span>|<span data-ttu-id="162c4-1494">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1494">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-1495">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1495">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="162c4-1496">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="162c4-1496">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="162c4-1497">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="162c4-1497">&lt;optional&gt;</span></span>|<span data-ttu-id="162c4-p191">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-p191">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="162c4-p192">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-p192">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="162c4-1502">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1502">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="162c4-1503">function</span><span class="sxs-lookup"><span data-stu-id="162c4-1503">function</span></span>||<span data-ttu-id="162c4-1504">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="162c4-1504">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="162c4-1505">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1505">Requirements</span></span>

|<span data-ttu-id="162c4-1506">要件</span><span class="sxs-lookup"><span data-stu-id="162c4-1506">Requirement</span></span>|<span data-ttu-id="162c4-1507">値</span><span class="sxs-lookup"><span data-stu-id="162c4-1507">Value</span></span>|
|---|---|
|[<span data-ttu-id="162c4-1508">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="162c4-1508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="162c4-1509">1.2</span><span class="sxs-lookup"><span data-stu-id="162c4-1509">1.2</span></span>|
|[<span data-ttu-id="162c4-1510">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="162c4-1510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="162c4-1511">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="162c4-1511">ReadWriteItem</span></span>|
|[<span data-ttu-id="162c4-1512">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="162c4-1512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="162c4-1513">Compose</span><span class="sxs-lookup"><span data-stu-id="162c4-1513">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="162c4-1514">例</span><span class="sxs-lookup"><span data-stu-id="162c4-1514">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```