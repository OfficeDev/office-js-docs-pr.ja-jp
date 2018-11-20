
# <a name="item"></a><span data-ttu-id="c01fc-101">item</span><span class="sxs-lookup"><span data-stu-id="c01fc-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="c01fc-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="c01fc-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="c01fc-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-105">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-105">Requirements</span></span>

|<span data-ttu-id="c01fc-106">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-106">Requirement</span></span>|<span data-ttu-id="c01fc-107">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-109">1.0</span></span>|
|[<span data-ttu-id="c01fc-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="c01fc-111">Restricted</span></span>|
|[<span data-ttu-id="c01fc-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-113">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c01fc-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c01fc-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="c01fc-114">Members and methods</span></span>

| <span data-ttu-id="c01fc-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-115">Member</span></span> | <span data-ttu-id="c01fc-116">種類</span><span class="sxs-lookup"><span data-stu-id="c01fc-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c01fc-117">attachments</span><span class="sxs-lookup"><span data-stu-id="c01fc-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails) | <span data-ttu-id="c01fc-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-118">Member</span></span> |
| [<span data-ttu-id="c01fc-119">bcc</span><span class="sxs-lookup"><span data-stu-id="c01fc-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="c01fc-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-120">Member</span></span> |
| [<span data-ttu-id="c01fc-121">body</span><span class="sxs-lookup"><span data-stu-id="c01fc-121">body</span></span>](#body-bodyjavascriptapioutlook17officebody) | <span data-ttu-id="c01fc-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-122">Member</span></span> |
| [<span data-ttu-id="c01fc-123">cc</span><span class="sxs-lookup"><span data-stu-id="c01fc-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="c01fc-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-124">Member</span></span> |
| [<span data-ttu-id="c01fc-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="c01fc-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="c01fc-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-126">Member</span></span> |
| [<span data-ttu-id="c01fc-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="c01fc-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="c01fc-128">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-128">Member</span></span> |
| [<span data-ttu-id="c01fc-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="c01fc-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="c01fc-130">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-130">Member</span></span> |
| [<span data-ttu-id="c01fc-131">end</span><span class="sxs-lookup"><span data-stu-id="c01fc-131">end</span></span>](#end-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="c01fc-132">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-132">Member</span></span> |
| [<span data-ttu-id="c01fc-133">from</span><span class="sxs-lookup"><span data-stu-id="c01fc-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) | <span data-ttu-id="c01fc-134">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-134">Member</span></span> |
| [<span data-ttu-id="c01fc-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="c01fc-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="c01fc-136">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-136">Member</span></span> |
| [<span data-ttu-id="c01fc-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="c01fc-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="c01fc-138">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-138">Member</span></span> |
| [<span data-ttu-id="c01fc-139">itemId</span><span class="sxs-lookup"><span data-stu-id="c01fc-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="c01fc-140">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-140">Member</span></span> |
| [<span data-ttu-id="c01fc-141">itemType</span><span class="sxs-lookup"><span data-stu-id="c01fc-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) | <span data-ttu-id="c01fc-142">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-142">Member</span></span> |
| [<span data-ttu-id="c01fc-143">location</span><span class="sxs-lookup"><span data-stu-id="c01fc-143">location</span></span>](#location-stringlocationjavascriptapioutlook17officelocation) | <span data-ttu-id="c01fc-144">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-144">Member</span></span> |
| [<span data-ttu-id="c01fc-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="c01fc-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="c01fc-146">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-146">Member</span></span> |
| [<span data-ttu-id="c01fc-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="c01fc-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages) | <span data-ttu-id="c01fc-148">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-148">Member</span></span> |
| [<span data-ttu-id="c01fc-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="c01fc-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="c01fc-150">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-150">Member</span></span> |
| [<span data-ttu-id="c01fc-151">organizer</span><span class="sxs-lookup"><span data-stu-id="c01fc-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) | <span data-ttu-id="c01fc-152">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-152">Member</span></span> |
| [<span data-ttu-id="c01fc-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="c01fc-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence) | <span data-ttu-id="c01fc-154">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-154">Member</span></span> |
| [<span data-ttu-id="c01fc-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="c01fc-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="c01fc-156">Member</span><span class="sxs-lookup"><span data-stu-id="c01fc-156">Member</span></span> |
| [<span data-ttu-id="c01fc-157">sender</span><span class="sxs-lookup"><span data-stu-id="c01fc-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) | <span data-ttu-id="c01fc-158">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-158">Member</span></span> |
| [<span data-ttu-id="c01fc-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="c01fc-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="c01fc-160">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-160">Member</span></span> |
| [<span data-ttu-id="c01fc-161">start</span><span class="sxs-lookup"><span data-stu-id="c01fc-161">start</span></span>](#start-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="c01fc-162">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-162">Member</span></span> |
| [<span data-ttu-id="c01fc-163">subject</span><span class="sxs-lookup"><span data-stu-id="c01fc-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlook17officesubject) | <span data-ttu-id="c01fc-164">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-164">Member</span></span> |
| [<span data-ttu-id="c01fc-165">to</span><span class="sxs-lookup"><span data-stu-id="c01fc-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="c01fc-166">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-166">Member</span></span> |
| [<span data-ttu-id="c01fc-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c01fc-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="c01fc-168">メソッド</span><span class="sxs-lookup"><span data-stu-id="c01fc-168">Method</span></span> |
| [<span data-ttu-id="c01fc-169">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c01fc-169">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="c01fc-170">メソッド</span><span class="sxs-lookup"><span data-stu-id="c01fc-170">Method</span></span> |
| [<span data-ttu-id="c01fc-171">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c01fc-171">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="c01fc-172">メソッド</span><span class="sxs-lookup"><span data-stu-id="c01fc-172">Method</span></span> |
| [<span data-ttu-id="c01fc-173">close</span><span class="sxs-lookup"><span data-stu-id="c01fc-173">close</span></span>](#close) | <span data-ttu-id="c01fc-174">メソッド</span><span class="sxs-lookup"><span data-stu-id="c01fc-174">Method</span></span> |
| [<span data-ttu-id="c01fc-175">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="c01fc-175">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="c01fc-176">メソッド</span><span class="sxs-lookup"><span data-stu-id="c01fc-176">Method</span></span> |
| [<span data-ttu-id="c01fc-177">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="c01fc-177">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="c01fc-178">メソッド</span><span class="sxs-lookup"><span data-stu-id="c01fc-178">Method</span></span> |
| [<span data-ttu-id="c01fc-179">getEntities</span><span class="sxs-lookup"><span data-stu-id="c01fc-179">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="c01fc-180">メソッド</span><span class="sxs-lookup"><span data-stu-id="c01fc-180">Method</span></span> |
| [<span data-ttu-id="c01fc-181">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="c01fc-181">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="c01fc-182">メソッド</span><span class="sxs-lookup"><span data-stu-id="c01fc-182">Method</span></span> |
| [<span data-ttu-id="c01fc-183">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="c01fc-183">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="c01fc-184">メソッド</span><span class="sxs-lookup"><span data-stu-id="c01fc-184">Method</span></span> |
| [<span data-ttu-id="c01fc-185">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="c01fc-185">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="c01fc-186">メソッド</span><span class="sxs-lookup"><span data-stu-id="c01fc-186">Method</span></span> |
| [<span data-ttu-id="c01fc-187">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="c01fc-187">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="c01fc-188">メソッド</span><span class="sxs-lookup"><span data-stu-id="c01fc-188">Method</span></span> |
| [<span data-ttu-id="c01fc-189">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c01fc-189">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="c01fc-190">メソッド</span><span class="sxs-lookup"><span data-stu-id="c01fc-190">Method</span></span> |
| [<span data-ttu-id="c01fc-191">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="c01fc-191">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="c01fc-192">メソッド</span><span class="sxs-lookup"><span data-stu-id="c01fc-192">Method</span></span> |
| [<span data-ttu-id="c01fc-193">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="c01fc-193">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="c01fc-194">メソッド</span><span class="sxs-lookup"><span data-stu-id="c01fc-194">Method</span></span> |
| [<span data-ttu-id="c01fc-195">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="c01fc-195">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="c01fc-196">メソッド</span><span class="sxs-lookup"><span data-stu-id="c01fc-196">Method</span></span> |
| [<span data-ttu-id="c01fc-197">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c01fc-197">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="c01fc-198">メソッド</span><span class="sxs-lookup"><span data-stu-id="c01fc-198">Method</span></span> |
| [<span data-ttu-id="c01fc-199">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c01fc-199">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="c01fc-200">メソッド</span><span class="sxs-lookup"><span data-stu-id="c01fc-200">Method</span></span> |
| [<span data-ttu-id="c01fc-201">saveAsync</span><span class="sxs-lookup"><span data-stu-id="c01fc-201">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="c01fc-202">メソッド</span><span class="sxs-lookup"><span data-stu-id="c01fc-202">Method</span></span> |
| [<span data-ttu-id="c01fc-203">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c01fc-203">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="c01fc-204">メソッド</span><span class="sxs-lookup"><span data-stu-id="c01fc-204">Method</span></span> |

### <a name="example"></a><span data-ttu-id="c01fc-205">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-205">Example</span></span>

<span data-ttu-id="c01fc-206">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c01fc-206">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
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

### <a name="members"></a><span data-ttu-id="c01fc-207">メンバー</span><span class="sxs-lookup"><span data-stu-id="c01fc-207">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="c01fc-208">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c01fc-208">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="c01fc-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c01fc-211">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="c01fc-211">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="c01fc-212">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c01fc-212">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-213">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-213">Type:</span></span>

*   <span data-ttu-id="c01fc-214">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c01fc-214">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-215">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-215">Requirements</span></span>

|<span data-ttu-id="c01fc-216">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-216">Requirement</span></span>|<span data-ttu-id="c01fc-217">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-218">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-219">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-219">1.0</span></span>|
|[<span data-ttu-id="c01fc-220">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-220">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-221">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-222">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-222">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-223">Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-223">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-224">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-224">Example</span></span>

<span data-ttu-id="c01fc-225">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-225">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
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

####  <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="c01fc-226">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c01fc-226">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="c01fc-227">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-227">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="c01fc-228">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c01fc-228">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-229">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-229">Type:</span></span>

*   [<span data-ttu-id="c01fc-230">Recipients</span><span class="sxs-lookup"><span data-stu-id="c01fc-230">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="c01fc-231">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-231">Requirements</span></span>

|<span data-ttu-id="c01fc-232">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-232">Requirement</span></span>|<span data-ttu-id="c01fc-233">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-234">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-234">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-235">1.1</span><span class="sxs-lookup"><span data-stu-id="c01fc-235">1.1</span></span>|
|[<span data-ttu-id="c01fc-236">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-236">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-237">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-237">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-238">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-238">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-239">Compose</span><span class="sxs-lookup"><span data-stu-id="c01fc-239">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-240">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-240">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="c01fc-241">body :[Body](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="c01fc-241">body :[Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="c01fc-242">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-242">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-243">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-243">Type:</span></span>

*   [<span data-ttu-id="c01fc-244">Body</span><span class="sxs-lookup"><span data-stu-id="c01fc-244">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="c01fc-245">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-245">Requirements</span></span>

|<span data-ttu-id="c01fc-246">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-246">Requirement</span></span>|<span data-ttu-id="c01fc-247">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-248">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-248">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-249">1.1</span><span class="sxs-lookup"><span data-stu-id="c01fc-249">1.1</span></span>|
|[<span data-ttu-id="c01fc-250">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-250">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-251">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-251">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-252">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-252">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-253">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-253">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="c01fc-254">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c01fc-254">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="c01fc-255">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-255">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="c01fc-256">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c01fc-256">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c01fc-257">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c01fc-257">Read mode</span></span>

<span data-ttu-id="c01fc-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c01fc-260">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c01fc-260">Compose mode</span></span>

<span data-ttu-id="c01fc-261">`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-261">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-262">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-262">Type:</span></span>

*   <span data-ttu-id="c01fc-263">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c01fc-263">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-264">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-264">Requirements</span></span>

|<span data-ttu-id="c01fc-265">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-265">Requirement</span></span>|<span data-ttu-id="c01fc-266">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-267">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-268">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-268">1.0</span></span>|
|[<span data-ttu-id="c01fc-269">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-269">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-270">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-271">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-271">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-272">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-272">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-273">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-273">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="c01fc-274">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="c01fc-274">(nullable) conversationId :String</span></span>

<span data-ttu-id="c01fc-275">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-275">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="c01fc-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="c01fc-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-280">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-280">Type:</span></span>

*   <span data-ttu-id="c01fc-281">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-281">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-282">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-282">Requirements</span></span>

|<span data-ttu-id="c01fc-283">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-283">Requirement</span></span>|<span data-ttu-id="c01fc-284">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-285">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-285">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-286">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-286">1.0</span></span>|
|[<span data-ttu-id="c01fc-287">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-287">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-288">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-289">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-289">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-290">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-290">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="c01fc-291">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="c01fc-291">dateTimeCreated :Date</span></span>

<span data-ttu-id="c01fc-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-294">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-294">Type:</span></span>

*   <span data-ttu-id="c01fc-295">日付</span><span class="sxs-lookup"><span data-stu-id="c01fc-295">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-296">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-296">Requirements</span></span>

|<span data-ttu-id="c01fc-297">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-297">Requirement</span></span>|<span data-ttu-id="c01fc-298">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-298">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-299">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-299">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-300">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-300">1.0</span></span>|
|[<span data-ttu-id="c01fc-301">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-301">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-302">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-302">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-303">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-303">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-304">Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-304">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-305">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-305">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="c01fc-306">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="c01fc-306">dateTimeModified :Date</span></span>

<span data-ttu-id="c01fc-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c01fc-309">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c01fc-309">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-310">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-310">Type:</span></span>

*   <span data-ttu-id="c01fc-311">日付</span><span class="sxs-lookup"><span data-stu-id="c01fc-311">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-312">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-312">Requirements</span></span>

|<span data-ttu-id="c01fc-313">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-313">Requirement</span></span>|<span data-ttu-id="c01fc-314">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-314">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-315">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-315">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-316">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-316">1.0</span></span>|
|[<span data-ttu-id="c01fc-317">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-317">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-318">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-318">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-319">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-319">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-320">Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-320">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-321">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-321">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="c01fc-322">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="c01fc-322">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="c01fc-323">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-323">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="c01fc-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c01fc-326">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c01fc-326">Read mode</span></span>

<span data-ttu-id="c01fc-327">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-327">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c01fc-328">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c01fc-328">Compose mode</span></span>

<span data-ttu-id="c01fc-329">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-329">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="c01fc-330">[`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c01fc-330">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-331">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-331">Type:</span></span>

*   <span data-ttu-id="c01fc-332">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="c01fc-332">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-333">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-333">Requirements</span></span>

|<span data-ttu-id="c01fc-334">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-334">Requirement</span></span>|<span data-ttu-id="c01fc-335">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-336">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-336">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-337">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-337">1.0</span></span>|
|[<span data-ttu-id="c01fc-338">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-339">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-340">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-341">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-342">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-342">Example</span></span>

<span data-ttu-id="c01fc-343">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-343">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="c01fc-344">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="c01fc-344">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="c01fc-345">メッセージの送信者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-345">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="c01fc-p112">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c01fc-348">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="c01fc-348">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c01fc-349">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c01fc-349">Read mode</span></span>

<span data-ttu-id="c01fc-350">`from` プロパティは `EmailAddressDetails` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-350">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="c01fc-351">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c01fc-351">Compose mode</span></span>

<span data-ttu-id="c01fc-352">`from` プロパティは From 値を取得するメソッドを提供する `From` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-352">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c01fc-353">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-353">Type:</span></span>

*   <span data-ttu-id="c01fc-354">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="c01fc-354">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-355">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-355">Requirements</span></span>

|<span data-ttu-id="c01fc-356">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-356">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="c01fc-357">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-358">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-358">1.0</span></span>|<span data-ttu-id="c01fc-359">1.7</span><span class="sxs-lookup"><span data-stu-id="c01fc-359">-17</span></span>|
|[<span data-ttu-id="c01fc-360">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-361">ReadItem</span></span>|<span data-ttu-id="c01fc-362">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-362">ReadWriteItem</span></span>|
|[<span data-ttu-id="c01fc-363">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-364">Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-364">Read</span></span>|<span data-ttu-id="c01fc-365">Compose</span><span class="sxs-lookup"><span data-stu-id="c01fc-365">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="c01fc-366">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="c01fc-366">internetMessageId :String</span></span>

<span data-ttu-id="c01fc-p113">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-369">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-369">Type:</span></span>

*   <span data-ttu-id="c01fc-370">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-370">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-371">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-371">Requirements</span></span>

|<span data-ttu-id="c01fc-372">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-372">Requirement</span></span>|<span data-ttu-id="c01fc-373">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-373">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-374">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-374">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-375">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-375">1.0</span></span>|
|[<span data-ttu-id="c01fc-376">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-376">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-377">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-378">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-378">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-379">Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-379">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-380">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-380">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="c01fc-381">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="c01fc-381">itemClass :String</span></span>

<span data-ttu-id="c01fc-p114">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="c01fc-p115">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="c01fc-386">型</span><span class="sxs-lookup"><span data-stu-id="c01fc-386">Type</span></span>|<span data-ttu-id="c01fc-387">説明</span><span class="sxs-lookup"><span data-stu-id="c01fc-387">Description</span></span>|<span data-ttu-id="c01fc-388">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="c01fc-388">item class</span></span>|
|---|---|---|
|<span data-ttu-id="c01fc-389">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="c01fc-389">Appointment items</span></span>|<span data-ttu-id="c01fc-390">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c01fc-390">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="c01fc-391">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="c01fc-391">Message items</span></span>|<span data-ttu-id="c01fc-392">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-392">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="c01fc-393">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-393">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-394">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-394">Type:</span></span>

*   <span data-ttu-id="c01fc-395">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-396">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-396">Requirements</span></span>

|<span data-ttu-id="c01fc-397">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-397">Requirement</span></span>|<span data-ttu-id="c01fc-398">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-399">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-399">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-400">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-400">1.0</span></span>|
|[<span data-ttu-id="c01fc-401">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-401">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-402">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-403">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-403">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-404">Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-405">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-405">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="c01fc-406">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="c01fc-406">(nullable) itemId :String</span></span>

<span data-ttu-id="c01fc-p116">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c01fc-409">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="c01fc-409">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="c01fc-410">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="c01fc-410">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="c01fc-411">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c01fc-411">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="c01fc-412">詳細については、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c01fc-412">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="c01fc-p118">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-415">種類:</span><span class="sxs-lookup"><span data-stu-id="c01fc-415">Type:</span></span>

*   <span data-ttu-id="c01fc-416">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-416">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-417">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-417">Requirements</span></span>

|<span data-ttu-id="c01fc-418">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-418">Requirement</span></span>|<span data-ttu-id="c01fc-419">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-420">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-420">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-421">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-421">1.0</span></span>|
|[<span data-ttu-id="c01fc-422">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-422">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-423">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-424">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-424">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-425">Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-425">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-426">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-426">Example</span></span>

<span data-ttu-id="c01fc-p119">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="c01fc-429">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="c01fc-429">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="c01fc-430">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-430">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="c01fc-431">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="c01fc-431">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-432">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-432">Type:</span></span>

*   [<span data-ttu-id="c01fc-433">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="c01fc-433">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="c01fc-434">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-434">Requirements</span></span>

|<span data-ttu-id="c01fc-435">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-435">Requirement</span></span>|<span data-ttu-id="c01fc-436">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-436">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-437">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-437">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-438">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-438">1.0</span></span>|
|[<span data-ttu-id="c01fc-439">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-439">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-440">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-440">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-441">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-441">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-442">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-442">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-443">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-443">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="c01fc-444">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="c01fc-444">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="c01fc-445">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-445">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c01fc-446">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c01fc-446">Read mode</span></span>

<span data-ttu-id="c01fc-447">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-447">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c01fc-448">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c01fc-448">Compose mode</span></span>

<span data-ttu-id="c01fc-449">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-449">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-450">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-450">Type:</span></span>

*   <span data-ttu-id="c01fc-451">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="c01fc-451">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-452">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-452">Requirements</span></span>

|<span data-ttu-id="c01fc-453">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-453">Requirement</span></span>|<span data-ttu-id="c01fc-454">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-455">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-455">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-456">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-456">1.0</span></span>|
|[<span data-ttu-id="c01fc-457">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-457">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-458">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-459">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-459">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-460">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-460">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-461">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-461">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="c01fc-462">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="c01fc-462">normalizedSubject :String</span></span>

<span data-ttu-id="c01fc-p120">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="c01fc-p121">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-467">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-467">Type:</span></span>

*   <span data-ttu-id="c01fc-468">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-468">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-469">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-469">Requirements</span></span>

|<span data-ttu-id="c01fc-470">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-470">Requirement</span></span>|<span data-ttu-id="c01fc-471">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-472">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-473">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-473">1.0</span></span>|
|[<span data-ttu-id="c01fc-474">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-474">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-475">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-476">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-476">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-477">Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-477">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-478">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-478">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="c01fc-479">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="c01fc-479">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="c01fc-480">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-480">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-481">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-481">Type:</span></span>

*   [<span data-ttu-id="c01fc-482">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="c01fc-482">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="c01fc-483">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-483">Requirements</span></span>

|<span data-ttu-id="c01fc-484">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-484">Requirement</span></span>|<span data-ttu-id="c01fc-485">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-485">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-486">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-486">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-487">1.3</span><span class="sxs-lookup"><span data-stu-id="c01fc-487">1.3</span></span>|
|[<span data-ttu-id="c01fc-488">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-488">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-489">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-489">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-490">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-490">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-491">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-491">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="c01fc-492">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c01fc-492">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="c01fc-493">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-493">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="c01fc-494">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c01fc-494">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c01fc-495">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c01fc-495">Read mode</span></span>

<span data-ttu-id="c01fc-496">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-496">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c01fc-497">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c01fc-497">Compose mode</span></span>

<span data-ttu-id="c01fc-498">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-498">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-499">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-499">Type:</span></span>

*   <span data-ttu-id="c01fc-500">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c01fc-500">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-501">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-501">Requirements</span></span>

|<span data-ttu-id="c01fc-502">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-502">Requirement</span></span>|<span data-ttu-id="c01fc-503">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-503">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-504">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-504">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-505">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-505">1.0</span></span>|
|[<span data-ttu-id="c01fc-506">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-506">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-507">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-507">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-508">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-508">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-509">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-509">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-510">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-510">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="c01fc-511">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="c01fc-511">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="c01fc-512">指定の会議の開催者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-512">Gets the email address of the meeting organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c01fc-513">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c01fc-513">Read mode</span></span>

<span data-ttu-id="c01fc-514">`organizer` プロパティは、会議開催者を表す [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-514">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c01fc-515">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c01fc-515">Compose mode</span></span>

<span data-ttu-id="c01fc-516">`organizer` プロパティは Organizer 値を取得するメソッドを提供する [Organizer](/javascript/api/outlook_1_7/office.organizer) オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-516">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-517">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-517">Type:</span></span>

*   <span data-ttu-id="c01fc-518">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="c01fc-518">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-519">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-519">Requirements</span></span>

|<span data-ttu-id="c01fc-520">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-520">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="c01fc-521">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-522">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-522">1.0</span></span>|<span data-ttu-id="c01fc-523">1.7</span><span class="sxs-lookup"><span data-stu-id="c01fc-523">-17</span></span>|
|[<span data-ttu-id="c01fc-524">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-525">ReadItem</span></span>|<span data-ttu-id="c01fc-526">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-526">ReadWriteItem</span></span>|
|[<span data-ttu-id="c01fc-527">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-527">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-528">Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-528">Read</span></span>|<span data-ttu-id="c01fc-529">Compose</span><span class="sxs-lookup"><span data-stu-id="c01fc-529">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-530">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-530">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="c01fc-531">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="c01fc-531">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="c01fc-532">予定の定期的なパターンを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-532">Gets or sets the location of an appointment.</span></span> <span data-ttu-id="c01fc-533">会議出席依頼の定期的なパターンを取得します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-533">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="c01fc-534">予定アイテムの閲覧モードと新規作成モード。</span><span class="sxs-lookup"><span data-stu-id="c01fc-534">Read and compose modes for appointment items.</span></span> <span data-ttu-id="c01fc-535">会議出席依頼アイテムの閲覧モード。</span><span class="sxs-lookup"><span data-stu-id="c01fc-535">Read mode for meeting request items.</span></span>

<span data-ttu-id="c01fc-536">`recurrence` プロパティは、アイテムがシリーズか、シリーズに含まれるインスタンスの場合、定期的な予定または会議出席依頼に対して [recurrence](/javascript/api/outlook_1_7/office.recurrence) オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-536">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="c01fc-537">`null` は、単発の予定および単発の予定の会議出席依頼に対して返されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-537">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="c01fc-538">`undefined` は、会議出席依頼ではないメッセージに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-538">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="c01fc-539">注: 会議出席依頼の `itemClass` 値は IPM.Schedule.Meeting.Request です。</span><span class="sxs-lookup"><span data-stu-id="c01fc-539">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="c01fc-540">注: recurrence オブジェクトが `null` の場合、オブジェクトがシリーズの一部ではなく、1 つの単発の予定または 1 つの単発の予定の会議出席依頼であることを示します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-540">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-541">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-541">Type:</span></span>

* [<span data-ttu-id="c01fc-542">Recurrence</span><span class="sxs-lookup"><span data-stu-id="c01fc-542">recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="c01fc-543">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-543">Requirement</span></span>|<span data-ttu-id="c01fc-544">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-545">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-546">1.7</span><span class="sxs-lookup"><span data-stu-id="c01fc-546">-17</span></span>|
|[<span data-ttu-id="c01fc-547">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-547">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-548">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-549">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-549">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-550">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-550">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="c01fc-551">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c01fc-551">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="c01fc-552">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-552">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="c01fc-553">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c01fc-553">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c01fc-554">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c01fc-554">Read mode</span></span>

<span data-ttu-id="c01fc-555">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-555">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c01fc-556">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c01fc-556">Compose mode</span></span>

<span data-ttu-id="c01fc-557">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-557">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-558">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-558">Type:</span></span>

*   <span data-ttu-id="c01fc-559">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c01fc-559">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-560">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-560">Requirements</span></span>

|<span data-ttu-id="c01fc-561">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-561">Requirement</span></span>|<span data-ttu-id="c01fc-562">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-563">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-564">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-564">1.0</span></span>|
|[<span data-ttu-id="c01fc-565">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-565">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-566">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-567">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-567">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-568">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-568">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-569">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-569">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="c01fc-570">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c01fc-570">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="c01fc-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="c01fc-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c01fc-575">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="c01fc-575">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-576">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-576">Type:</span></span>

*   [<span data-ttu-id="c01fc-577">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c01fc-577">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="c01fc-578">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-578">Requirements</span></span>

|<span data-ttu-id="c01fc-579">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-579">Requirement</span></span>|<span data-ttu-id="c01fc-580">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-581">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-581">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-582">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-582">1.0</span></span>|
|[<span data-ttu-id="c01fc-583">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-583">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-584">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-585">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-585">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-586">Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-586">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-587">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-587">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="c01fc-588">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="c01fc-588">(nullable) seriesId :String</span></span>

<span data-ttu-id="c01fc-589">あるインスタンスが属するシリーズの ID を取得します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-589">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="c01fc-590">OWA と Outlook では、`seriesId` はこのアイテムが属する親 (シリーズ) アイテムの Exchange Web Services (EWS) ID を返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-590">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="c01fc-591">ただし、iOS と Android の場合、`seriesId` は親アイテムの REST ID を返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-591">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="c01fc-592">`seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="c01fc-592">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="c01fc-593">`seriesId` プロパティは、Outlook REST API で使用される Outlook ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="c01fc-593">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="c01fc-594">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c01fc-594">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="c01fc-595">詳細については、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c01fc-595">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="c01fc-596">`seriesId` プロパティは、単発の予定、シリーズ アイテム、会議出席依頼など、親アイテムを持たないアイテムに対して `null` を返し、会議出席依頼ではないその他のアイテムに対して `undefined` を返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-596">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-597">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-597">Type:</span></span>

* <span data-ttu-id="c01fc-598">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-598">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-599">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-599">Requirements</span></span>

|<span data-ttu-id="c01fc-600">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-600">Requirement</span></span>|<span data-ttu-id="c01fc-601">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-601">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-602">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-602">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-603">1.7</span><span class="sxs-lookup"><span data-stu-id="c01fc-603">-17</span></span>|
|[<span data-ttu-id="c01fc-604">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-604">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-605">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-605">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-606">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-606">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-607">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-607">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-608">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-608">Example</span></span>

```js
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="c01fc-609">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="c01fc-609">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="c01fc-610">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-610">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="c01fc-p130">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c01fc-613">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c01fc-613">Read mode</span></span>

<span data-ttu-id="c01fc-614">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-614">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c01fc-615">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c01fc-615">Compose mode</span></span>

<span data-ttu-id="c01fc-616">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-616">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="c01fc-617">[`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c01fc-617">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-618">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-618">Type:</span></span>

*   <span data-ttu-id="c01fc-619">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="c01fc-619">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-620">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-620">Requirements</span></span>

|<span data-ttu-id="c01fc-621">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-621">Requirement</span></span>|<span data-ttu-id="c01fc-622">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-622">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-623">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-623">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-624">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-624">1.0</span></span>|
|[<span data-ttu-id="c01fc-625">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-625">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-626">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-626">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-627">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-627">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-628">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-628">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-629">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-629">Example</span></span>

<span data-ttu-id="c01fc-630">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-630">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

####  <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="c01fc-631">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c01fc-631">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="c01fc-632">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-632">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="c01fc-633">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-633">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c01fc-634">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c01fc-634">Read mode</span></span>

<span data-ttu-id="c01fc-p131">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="c01fc-637">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c01fc-637">Compose mode</span></span>

<span data-ttu-id="c01fc-638">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-638">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c01fc-639">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-639">Type:</span></span>

*   <span data-ttu-id="c01fc-640">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c01fc-640">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-641">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-641">Requirements</span></span>

|<span data-ttu-id="c01fc-642">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-642">Requirement</span></span>|<span data-ttu-id="c01fc-643">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-644">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-644">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-645">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-645">1.0</span></span>|
|[<span data-ttu-id="c01fc-646">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-646">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-647">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-647">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-648">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-648">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-649">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-649">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="c01fc-650">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c01fc-650">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="c01fc-651">メッセージの**宛先**行の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-651">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="c01fc-652">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c01fc-652">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c01fc-653">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c01fc-653">Read mode</span></span>

<span data-ttu-id="c01fc-p133">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c01fc-656">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c01fc-656">Compose mode</span></span>

<span data-ttu-id="c01fc-657">`to` プロパティは、メッセージの**宛先**行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-657">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="c01fc-658">型:</span><span class="sxs-lookup"><span data-stu-id="c01fc-658">Type:</span></span>

*   <span data-ttu-id="c01fc-659">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c01fc-659">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-660">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-660">Requirements</span></span>

|<span data-ttu-id="c01fc-661">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-661">Requirement</span></span>|<span data-ttu-id="c01fc-662">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-662">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-663">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-663">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-664">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-664">1.0</span></span>|
|[<span data-ttu-id="c01fc-665">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-665">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-666">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-666">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-667">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-667">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-668">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-668">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-669">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-669">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="c01fc-670">メソッド</span><span class="sxs-lookup"><span data-stu-id="c01fc-670">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="c01fc-671">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c01fc-671">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c01fc-672">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-672">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="c01fc-673">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-673">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="c01fc-674">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-674">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c01fc-675">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c01fc-675">Parameters:</span></span>
|<span data-ttu-id="c01fc-676">名前</span><span class="sxs-lookup"><span data-stu-id="c01fc-676">Name</span></span>|<span data-ttu-id="c01fc-677">型</span><span class="sxs-lookup"><span data-stu-id="c01fc-677">Type</span></span>|<span data-ttu-id="c01fc-678">属性</span><span class="sxs-lookup"><span data-stu-id="c01fc-678">Attributes</span></span>|<span data-ttu-id="c01fc-679">説明</span><span class="sxs-lookup"><span data-stu-id="c01fc-679">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="c01fc-680">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-680">String</span></span>||<span data-ttu-id="c01fc-p134">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="c01fc-683">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-683">String</span></span>||<span data-ttu-id="c01fc-p135">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="c01fc-686">Object</span><span class="sxs-lookup"><span data-stu-id="c01fc-686">Object</span></span>|<span data-ttu-id="c01fc-687">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-687">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-688">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c01fc-688">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c01fc-689">Object</span><span class="sxs-lookup"><span data-stu-id="c01fc-689">Object</span></span>|<span data-ttu-id="c01fc-690">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-690">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-691">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-691">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="c01fc-692">Boolean</span><span class="sxs-lookup"><span data-stu-id="c01fc-692">Boolean</span></span>|<span data-ttu-id="c01fc-693">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-693">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-694">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-694">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="c01fc-695">function</span><span class="sxs-lookup"><span data-stu-id="c01fc-695">function</span></span>|<span data-ttu-id="c01fc-696">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-696">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-697">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-697">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c01fc-698">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-698">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c01fc-699">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-699">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c01fc-700">エラー</span><span class="sxs-lookup"><span data-stu-id="c01fc-700">Errors</span></span>

|<span data-ttu-id="c01fc-701">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c01fc-701">Error code</span></span>|<span data-ttu-id="c01fc-702">説明</span><span class="sxs-lookup"><span data-stu-id="c01fc-702">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="c01fc-703">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="c01fc-703">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="c01fc-704">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="c01fc-704">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="c01fc-705">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-705">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c01fc-706">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-706">Requirements</span></span>

|<span data-ttu-id="c01fc-707">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-707">Requirement</span></span>|<span data-ttu-id="c01fc-708">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-708">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-709">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-709">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-710">1.1</span><span class="sxs-lookup"><span data-stu-id="c01fc-710">1.1</span></span>|
|[<span data-ttu-id="c01fc-711">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-711">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-712">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-712">ReadWriteItem</span></span>|
|[<span data-ttu-id="c01fc-713">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-713">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-714">Compose</span><span class="sxs-lookup"><span data-stu-id="c01fc-714">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c01fc-715">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-715">Examples</span></span>

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

<span data-ttu-id="c01fc-716">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-716">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="c01fc-717">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c01fc-717">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="c01fc-718">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-718">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="c01fc-719">現在、サポートされているイベントの種類は `Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged`、`Office.EventType.RecurrenceChanged` です。</span><span class="sxs-lookup"><span data-stu-id="c01fc-719">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="c01fc-720">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c01fc-720">Parameters:</span></span>

| <span data-ttu-id="c01fc-721">名前</span><span class="sxs-lookup"><span data-stu-id="c01fc-721">Name</span></span> | <span data-ttu-id="c01fc-722">型</span><span class="sxs-lookup"><span data-stu-id="c01fc-722">Type</span></span> | <span data-ttu-id="c01fc-723">属性</span><span class="sxs-lookup"><span data-stu-id="c01fc-723">Attributes</span></span> | <span data-ttu-id="c01fc-724">説明</span><span class="sxs-lookup"><span data-stu-id="c01fc-724">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c01fc-725">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c01fc-725">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c01fc-726">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="c01fc-726">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="c01fc-727">Function</span><span class="sxs-lookup"><span data-stu-id="c01fc-727">Function</span></span> || <span data-ttu-id="c01fc-p136">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p136">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="c01fc-731">Object</span><span class="sxs-lookup"><span data-stu-id="c01fc-731">Object</span></span> | <span data-ttu-id="c01fc-732">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-732">&lt;optional&gt;</span></span> | <span data-ttu-id="c01fc-733">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c01fc-733">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c01fc-734">Object</span><span class="sxs-lookup"><span data-stu-id="c01fc-734">Object</span></span> | <span data-ttu-id="c01fc-735">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-735">&lt;optional&gt;</span></span> | <span data-ttu-id="c01fc-736">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-736">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c01fc-737">function</span><span class="sxs-lookup"><span data-stu-id="c01fc-737">function</span></span>| <span data-ttu-id="c01fc-738">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-738">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-739">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-739">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c01fc-740">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-740">Requirements</span></span>

|<span data-ttu-id="c01fc-741">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-741">Requirement</span></span>| <span data-ttu-id="c01fc-742">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-742">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-743">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-743">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c01fc-744">1.7</span><span class="sxs-lookup"><span data-stu-id="c01fc-744">-17</span></span> |
|[<span data-ttu-id="c01fc-745">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-745">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c01fc-746">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-746">ReadItem</span></span> |
|[<span data-ttu-id="c01fc-747">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-747">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c01fc-748">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-748">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="c01fc-749">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-749">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="c01fc-750">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c01fc-750">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c01fc-751">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-751">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="c01fc-p137">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p137">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="c01fc-755">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-755">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="c01fc-756">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、推奨されていません。</span><span class="sxs-lookup"><span data-stu-id="c01fc-756">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c01fc-757">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c01fc-757">Parameters:</span></span>

|<span data-ttu-id="c01fc-758">名前</span><span class="sxs-lookup"><span data-stu-id="c01fc-758">Name</span></span>|<span data-ttu-id="c01fc-759">型</span><span class="sxs-lookup"><span data-stu-id="c01fc-759">Type</span></span>|<span data-ttu-id="c01fc-760">属性</span><span class="sxs-lookup"><span data-stu-id="c01fc-760">Attributes</span></span>|<span data-ttu-id="c01fc-761">説明</span><span class="sxs-lookup"><span data-stu-id="c01fc-761">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="c01fc-762">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-762">String</span></span>||<span data-ttu-id="c01fc-p138">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p138">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="c01fc-765">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-765">String</span></span>||<span data-ttu-id="c01fc-p139">添付するアイテムの件名。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p139">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="c01fc-768">Object</span><span class="sxs-lookup"><span data-stu-id="c01fc-768">Object</span></span>|<span data-ttu-id="c01fc-769">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-769">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-770">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c01fc-770">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c01fc-771">Object</span><span class="sxs-lookup"><span data-stu-id="c01fc-771">Object</span></span>|<span data-ttu-id="c01fc-772">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-772">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-773">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-773">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c01fc-774">function</span><span class="sxs-lookup"><span data-stu-id="c01fc-774">function</span></span>|<span data-ttu-id="c01fc-775">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-775">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-776">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-776">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c01fc-777">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-777">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c01fc-778">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-778">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c01fc-779">エラー</span><span class="sxs-lookup"><span data-stu-id="c01fc-779">Errors</span></span>

|<span data-ttu-id="c01fc-780">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c01fc-780">Error code</span></span>|<span data-ttu-id="c01fc-781">説明</span><span class="sxs-lookup"><span data-stu-id="c01fc-781">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="c01fc-782">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-782">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c01fc-783">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-783">Requirements</span></span>

|<span data-ttu-id="c01fc-784">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-784">Requirement</span></span>|<span data-ttu-id="c01fc-785">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-785">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-786">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-786">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-787">1.1</span><span class="sxs-lookup"><span data-stu-id="c01fc-787">1.1</span></span>|
|[<span data-ttu-id="c01fc-788">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-788">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-789">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-789">ReadWriteItem</span></span>|
|[<span data-ttu-id="c01fc-790">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-790">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-791">Compose</span><span class="sxs-lookup"><span data-stu-id="c01fc-791">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-792">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-792">Example</span></span>

<span data-ttu-id="c01fc-793">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-793">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
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

####  <a name="close"></a><span data-ttu-id="c01fc-794">close()</span><span class="sxs-lookup"><span data-stu-id="c01fc-794">close()</span></span>

<span data-ttu-id="c01fc-795">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-795">Closes the current item that is being composed.</span></span>

<span data-ttu-id="c01fc-p140">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p140">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="c01fc-798">Outlook on the web では、アイテムが予定であり、`saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、またはキャンセルするようにユーザーは求められます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-798">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="c01fc-799">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="c01fc-799">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-800">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-800">Requirements</span></span>

|<span data-ttu-id="c01fc-801">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-801">Requirement</span></span>|<span data-ttu-id="c01fc-802">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-802">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-803">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-803">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-804">1.3</span><span class="sxs-lookup"><span data-stu-id="c01fc-804">1.3</span></span>|
|[<span data-ttu-id="c01fc-805">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-805">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-806">制限あり</span><span class="sxs-lookup"><span data-stu-id="c01fc-806">Restricted</span></span>|
|[<span data-ttu-id="c01fc-807">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-807">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-808">Compose</span><span class="sxs-lookup"><span data-stu-id="c01fc-808">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="c01fc-809">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="c01fc-809">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="c01fc-810">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-810">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c01fc-811">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c01fc-811">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c01fc-812">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-812">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c01fc-813">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="c01fc-813">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="c01fc-p141">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p141">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c01fc-817">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c01fc-817">Parameters:</span></span>

|<span data-ttu-id="c01fc-818">名前</span><span class="sxs-lookup"><span data-stu-id="c01fc-818">Name</span></span>|<span data-ttu-id="c01fc-819">型</span><span class="sxs-lookup"><span data-stu-id="c01fc-819">Type</span></span>|<span data-ttu-id="c01fc-820">属性</span><span class="sxs-lookup"><span data-stu-id="c01fc-820">Attributes</span></span>|<span data-ttu-id="c01fc-821">説明</span><span class="sxs-lookup"><span data-stu-id="c01fc-821">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="c01fc-822">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="c01fc-822">String &#124; Object</span></span>||<span data-ttu-id="c01fc-p142">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c01fc-825">**または**</span><span class="sxs-lookup"><span data-stu-id="c01fc-825">**OR**</span></span><br/><span data-ttu-id="c01fc-p143">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="c01fc-828">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-828">String</span></span>|<span data-ttu-id="c01fc-829">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-829">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-p144">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="c01fc-832">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-832">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="c01fc-833">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-833">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-834">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="c01fc-834">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="c01fc-835">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-835">String</span></span>||<span data-ttu-id="c01fc-p145">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p145">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="c01fc-838">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-838">String</span></span>||<span data-ttu-id="c01fc-839">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c01fc-839">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="c01fc-840">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-840">String</span></span>||<span data-ttu-id="c01fc-p146">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p146">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="c01fc-843">ブール値</span><span class="sxs-lookup"><span data-stu-id="c01fc-843">Boolean</span></span>||<span data-ttu-id="c01fc-p147">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p147">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="c01fc-846">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-846">String</span></span>||<span data-ttu-id="c01fc-p148">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p148">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="c01fc-850">function</span><span class="sxs-lookup"><span data-stu-id="c01fc-850">function</span></span>|<span data-ttu-id="c01fc-851">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-851">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-852">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-852">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c01fc-853">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-853">Requirements</span></span>

|<span data-ttu-id="c01fc-854">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-854">Requirement</span></span>|<span data-ttu-id="c01fc-855">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-856">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-856">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-857">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-857">1.0</span></span>|
|[<span data-ttu-id="c01fc-858">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-858">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-859">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-859">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-860">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-860">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-861">Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-861">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c01fc-862">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-862">Examples</span></span>

<span data-ttu-id="c01fc-863">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-863">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="c01fc-864">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-864">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="c01fc-865">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-865">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c01fc-866">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-866">Reply with a body and a file attachment.</span></span>

```js
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

<span data-ttu-id="c01fc-867">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-867">Reply with a body and an item attachment.</span></span>

```js
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

<span data-ttu-id="c01fc-868">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-868">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="c01fc-869">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="c01fc-869">displayReplyForm(formData)</span></span>

<span data-ttu-id="c01fc-870">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-870">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c01fc-871">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c01fc-871">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c01fc-872">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-872">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c01fc-873">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="c01fc-873">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="c01fc-p149">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p149">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c01fc-877">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c01fc-877">Parameters:</span></span>

|<span data-ttu-id="c01fc-878">名前</span><span class="sxs-lookup"><span data-stu-id="c01fc-878">Name</span></span>|<span data-ttu-id="c01fc-879">型</span><span class="sxs-lookup"><span data-stu-id="c01fc-879">Type</span></span>|<span data-ttu-id="c01fc-880">属性</span><span class="sxs-lookup"><span data-stu-id="c01fc-880">Attributes</span></span>|<span data-ttu-id="c01fc-881">説明</span><span class="sxs-lookup"><span data-stu-id="c01fc-881">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="c01fc-882">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="c01fc-882">String &#124; Object</span></span>||<span data-ttu-id="c01fc-p150">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p150">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c01fc-885">**または**</span><span class="sxs-lookup"><span data-stu-id="c01fc-885">**OR**</span></span><br/><span data-ttu-id="c01fc-p151">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p151">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="c01fc-888">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-888">String</span></span>|<span data-ttu-id="c01fc-889">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-889">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-p152">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="c01fc-892">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-892">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="c01fc-893">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-893">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-894">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="c01fc-894">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="c01fc-895">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-895">String</span></span>||<span data-ttu-id="c01fc-p153">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p153">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="c01fc-898">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-898">String</span></span>||<span data-ttu-id="c01fc-899">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c01fc-899">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="c01fc-900">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-900">String</span></span>||<span data-ttu-id="c01fc-p154">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p154">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="c01fc-903">ブール値</span><span class="sxs-lookup"><span data-stu-id="c01fc-903">Boolean</span></span>||<span data-ttu-id="c01fc-p155">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p155">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="c01fc-906">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-906">String</span></span>||<span data-ttu-id="c01fc-p156">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p156">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="c01fc-910">function</span><span class="sxs-lookup"><span data-stu-id="c01fc-910">function</span></span>|<span data-ttu-id="c01fc-911">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-911">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-912">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-912">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c01fc-913">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-913">Requirements</span></span>

|<span data-ttu-id="c01fc-914">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-914">Requirement</span></span>|<span data-ttu-id="c01fc-915">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-915">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-916">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-916">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-917">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-917">1.0</span></span>|
|[<span data-ttu-id="c01fc-918">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-918">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-919">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-919">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-920">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-920">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-921">Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-921">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c01fc-922">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-922">Examples</span></span>

<span data-ttu-id="c01fc-923">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-923">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="c01fc-924">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-924">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="c01fc-925">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-925">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c01fc-926">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-926">Reply with a body and a file attachment.</span></span>

```js
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

<span data-ttu-id="c01fc-927">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-927">Reply with a body and an item attachment.</span></span>

```js
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

<span data-ttu-id="c01fc-928">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-928">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
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

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="c01fc-929">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="c01fc-929">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="c01fc-930">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-930">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="c01fc-931">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c01fc-931">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-932">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-932">Requirements</span></span>

|<span data-ttu-id="c01fc-933">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-933">Requirement</span></span>|<span data-ttu-id="c01fc-934">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-935">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-936">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-936">1.0</span></span>|
|[<span data-ttu-id="c01fc-937">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-938">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-939">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-940">Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-940">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c01fc-941">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c01fc-941">Returns:</span></span>

<span data-ttu-id="c01fc-942">型:[Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="c01fc-942">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="c01fc-943">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-943">Example</span></span>

<span data-ttu-id="c01fc-944">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="c01fc-944">The following example accesses the contacts entities on the current item.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="c01fc-945">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="c01fc-945">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c01fc-946">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-946">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="c01fc-947">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c01fc-947">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c01fc-948">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c01fc-948">Parameters:</span></span>

|<span data-ttu-id="c01fc-949">名前</span><span class="sxs-lookup"><span data-stu-id="c01fc-949">Name</span></span>|<span data-ttu-id="c01fc-950">型</span><span class="sxs-lookup"><span data-stu-id="c01fc-950">Type</span></span>|<span data-ttu-id="c01fc-951">説明</span><span class="sxs-lookup"><span data-stu-id="c01fc-951">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="c01fc-952">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="c01fc-952">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="c01fc-953">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="c01fc-953">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c01fc-954">Requirements</span><span class="sxs-lookup"><span data-stu-id="c01fc-954">Requirements</span></span>

|<span data-ttu-id="c01fc-955">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-955">Requirement</span></span>|<span data-ttu-id="c01fc-956">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-957">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-957">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-958">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-958">1.0</span></span>|
|[<span data-ttu-id="c01fc-959">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-960">制限あり</span><span class="sxs-lookup"><span data-stu-id="c01fc-960">Restricted</span></span>|
|[<span data-ttu-id="c01fc-961">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-962">Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c01fc-963">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c01fc-963">Returns:</span></span>

<span data-ttu-id="c01fc-964">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-964">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="c01fc-965">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-965">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="c01fc-966">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="c01fc-966">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="c01fc-967">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="c01fc-967">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="c01fc-968">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="c01fc-968">Value of `entityType`</span></span>|<span data-ttu-id="c01fc-969">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="c01fc-969">Type of objects in returned array</span></span>|<span data-ttu-id="c01fc-970">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-970">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="c01fc-971">文字列</span><span class="sxs-lookup"><span data-stu-id="c01fc-971">String</span></span>|<span data-ttu-id="c01fc-972">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="c01fc-972">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="c01fc-973">連絡先</span><span class="sxs-lookup"><span data-stu-id="c01fc-973">Contact</span></span>|<span data-ttu-id="c01fc-974">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c01fc-974">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="c01fc-975">文字列</span><span class="sxs-lookup"><span data-stu-id="c01fc-975">String</span></span>|<span data-ttu-id="c01fc-976">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c01fc-976">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="c01fc-977">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="c01fc-977">MeetingSuggestion</span></span>|<span data-ttu-id="c01fc-978">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c01fc-978">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="c01fc-979">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="c01fc-979">PhoneNumber</span></span>|<span data-ttu-id="c01fc-980">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="c01fc-980">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="c01fc-981">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="c01fc-981">TaskSuggestion</span></span>|<span data-ttu-id="c01fc-982">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c01fc-982">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="c01fc-983">文字列</span><span class="sxs-lookup"><span data-stu-id="c01fc-983">String</span></span>|<span data-ttu-id="c01fc-984">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="c01fc-984">**Restricted**</span></span>|

<span data-ttu-id="c01fc-985">型:Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c01fc-985">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="c01fc-986">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-986">Example</span></span>

<span data-ttu-id="c01fc-987">次の例は、現在のアイテムの本文にある郵便アドレスを表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-987">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

```js
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="c01fc-988">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="c01fc-988">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c01fc-989">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-989">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c01fc-990">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c01fc-990">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c01fc-991">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-991">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c01fc-992">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c01fc-992">Parameters:</span></span>

|<span data-ttu-id="c01fc-993">名前</span><span class="sxs-lookup"><span data-stu-id="c01fc-993">Name</span></span>|<span data-ttu-id="c01fc-994">型</span><span class="sxs-lookup"><span data-stu-id="c01fc-994">Type</span></span>|<span data-ttu-id="c01fc-995">説明</span><span class="sxs-lookup"><span data-stu-id="c01fc-995">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="c01fc-996">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-996">String</span></span>|<span data-ttu-id="c01fc-997">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="c01fc-997">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c01fc-998">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-998">Requirements</span></span>

|<span data-ttu-id="c01fc-999">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-999">Requirement</span></span>|<span data-ttu-id="c01fc-1000">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-1000">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-1001">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-1001">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-1002">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-1002">1.0</span></span>|
|[<span data-ttu-id="c01fc-1003">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-1003">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-1004">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-1004">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-1005">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-1005">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-1006">Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-1006">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c01fc-1007">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c01fc-1007">Returns:</span></span>

<span data-ttu-id="c01fc-p158">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p158">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="c01fc-1010">型:Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c01fc-1010">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="c01fc-1011">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="c01fc-1011">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="c01fc-1012">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1012">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c01fc-1013">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1013">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c01fc-p159">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p159">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c01fc-1017">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1017">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c01fc-1018">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1018">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c01fc-p160">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p160">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-1022">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-1022">Requirements</span></span>

|<span data-ttu-id="c01fc-1023">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-1023">Requirement</span></span>|<span data-ttu-id="c01fc-1024">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-1024">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-1025">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-1025">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-1026">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-1026">1.0</span></span>|
|[<span data-ttu-id="c01fc-1027">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-1027">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-1028">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-1028">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-1029">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-1029">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-1030">Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-1030">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c01fc-1031">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c01fc-1031">Returns:</span></span>

<span data-ttu-id="c01fc-p161">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p161">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="c01fc-1034">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c01fc-1034">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c01fc-1035">Object</span><span class="sxs-lookup"><span data-stu-id="c01fc-1035">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c01fc-1036">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-1036">Example</span></span>

<span data-ttu-id="c01fc-1037">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1037">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="c01fc-1038">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="c01fc-1038">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="c01fc-1039">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定の正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1039">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c01fc-1040">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1040">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c01fc-1041">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1041">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="c01fc-p162">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c01fc-1044">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c01fc-1044">Parameters:</span></span>

|<span data-ttu-id="c01fc-1045">名前</span><span class="sxs-lookup"><span data-stu-id="c01fc-1045">Name</span></span>|<span data-ttu-id="c01fc-1046">型</span><span class="sxs-lookup"><span data-stu-id="c01fc-1046">Type</span></span>|<span data-ttu-id="c01fc-1047">説明</span><span class="sxs-lookup"><span data-stu-id="c01fc-1047">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="c01fc-1048">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-1048">String</span></span>|<span data-ttu-id="c01fc-1049">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1049">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c01fc-1050">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-1050">Requirements</span></span>

|<span data-ttu-id="c01fc-1051">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-1051">Requirement</span></span>|<span data-ttu-id="c01fc-1052">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-1052">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-1053">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-1053">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-1054">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-1054">1.0</span></span>|
|[<span data-ttu-id="c01fc-1055">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-1055">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-1056">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-1056">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-1057">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-1057">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-1058">Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-1058">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c01fc-1059">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c01fc-1059">Returns:</span></span>

<span data-ttu-id="c01fc-1060">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1060">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="c01fc-1061">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c01fc-1061">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c01fc-1062">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="c01fc-1062">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c01fc-1063">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-1063">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="c01fc-1064">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="c01fc-1064">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="c01fc-1065">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1065">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="c01fc-p163">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p163">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c01fc-1068">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c01fc-1068">Parameters:</span></span>

|<span data-ttu-id="c01fc-1069">名前</span><span class="sxs-lookup"><span data-stu-id="c01fc-1069">Name</span></span>|<span data-ttu-id="c01fc-1070">型</span><span class="sxs-lookup"><span data-stu-id="c01fc-1070">Type</span></span>|<span data-ttu-id="c01fc-1071">属性</span><span class="sxs-lookup"><span data-stu-id="c01fc-1071">Attributes</span></span>|<span data-ttu-id="c01fc-1072">説明</span><span class="sxs-lookup"><span data-stu-id="c01fc-1072">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="c01fc-1073">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c01fc-1073">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="c01fc-p164">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p164">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="c01fc-1077">Object</span><span class="sxs-lookup"><span data-stu-id="c01fc-1077">Object</span></span>|<span data-ttu-id="c01fc-1078">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-1078">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-1079">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1079">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c01fc-1080">Object</span><span class="sxs-lookup"><span data-stu-id="c01fc-1080">Object</span></span>|<span data-ttu-id="c01fc-1081">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-1081">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-1082">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1082">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c01fc-1083">function</span><span class="sxs-lookup"><span data-stu-id="c01fc-1083">function</span></span>||<span data-ttu-id="c01fc-1084">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1084">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c01fc-1085">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1085">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="c01fc-1086">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1086">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c01fc-1087">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-1087">Requirements</span></span>

|<span data-ttu-id="c01fc-1088">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-1088">Requirement</span></span>|<span data-ttu-id="c01fc-1089">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-1089">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-1090">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-1090">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-1091">1.2</span><span class="sxs-lookup"><span data-stu-id="c01fc-1091">1.2</span></span>|
|[<span data-ttu-id="c01fc-1092">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-1092">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-1093">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-1093">ReadWriteItem</span></span>|
|[<span data-ttu-id="c01fc-1094">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-1094">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-1095">Compose</span><span class="sxs-lookup"><span data-stu-id="c01fc-1095">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="c01fc-1096">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c01fc-1096">Returns:</span></span>

<span data-ttu-id="c01fc-1097">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1097">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="c01fc-1098">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c01fc-1098">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c01fc-1099">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-1099">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c01fc-1100">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-1100">Example</span></span>

```js
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

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="c01fc-1101">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="c01fc-1101">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="c01fc-p166">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p166">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="c01fc-1104">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1104">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-1105">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-1105">Requirements</span></span>

|<span data-ttu-id="c01fc-1106">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-1106">Requirement</span></span>|<span data-ttu-id="c01fc-1107">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-1107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-1108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-1108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-1109">1.6</span><span class="sxs-lookup"><span data-stu-id="c01fc-1109">-16</span></span>|
|[<span data-ttu-id="c01fc-1110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-1110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-1111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-1111">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-1112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-1112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-1113">Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-1113">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c01fc-1114">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c01fc-1114">Returns:</span></span>

<span data-ttu-id="c01fc-1115">型:[Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="c01fc-1115">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="c01fc-1116">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-1116">Example</span></span>

<span data-ttu-id="c01fc-1117">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1117">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="c01fc-1118">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="c01fc-1118">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="c01fc-p167">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p167">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="c01fc-1121">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1121">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c01fc-p168">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p168">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c01fc-1125">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1125">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c01fc-1126">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1126">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c01fc-p169">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c01fc-1130">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-1130">Requirements</span></span>

|<span data-ttu-id="c01fc-1131">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-1131">Requirement</span></span>|<span data-ttu-id="c01fc-1132">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-1132">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-1133">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-1133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-1134">1.6</span><span class="sxs-lookup"><span data-stu-id="c01fc-1134">-16</span></span>|
|[<span data-ttu-id="c01fc-1135">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-1135">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-1136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-1136">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-1137">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-1137">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-1138">Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-1138">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c01fc-1139">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c01fc-1139">Returns:</span></span>

<span data-ttu-id="c01fc-p170">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="c01fc-1142">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-1142">Example</span></span>

<span data-ttu-id="c01fc-1143">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1143">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="c01fc-1144">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c01fc-1144">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="c01fc-1145">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1145">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="c01fc-p171">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p171">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c01fc-1149">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c01fc-1149">Parameters:</span></span>

|<span data-ttu-id="c01fc-1150">名前</span><span class="sxs-lookup"><span data-stu-id="c01fc-1150">Name</span></span>|<span data-ttu-id="c01fc-1151">型</span><span class="sxs-lookup"><span data-stu-id="c01fc-1151">Type</span></span>|<span data-ttu-id="c01fc-1152">属性</span><span class="sxs-lookup"><span data-stu-id="c01fc-1152">Attributes</span></span>|<span data-ttu-id="c01fc-1153">説明</span><span class="sxs-lookup"><span data-stu-id="c01fc-1153">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="c01fc-1154">function</span><span class="sxs-lookup"><span data-stu-id="c01fc-1154">function</span></span>||<span data-ttu-id="c01fc-1155">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1155">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c01fc-1156">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1156">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="c01fc-1157">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、および削除し、カスタム プロパティに対する変更をサーバーに設定し直すために使用できます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1157">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="c01fc-1158">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c01fc-1158">Object</span></span>|<span data-ttu-id="c01fc-1159">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-1159">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-1160">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1160">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="c01fc-1161">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1161">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c01fc-1162">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-1162">Requirements</span></span>

|<span data-ttu-id="c01fc-1163">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-1163">Requirement</span></span>|<span data-ttu-id="c01fc-1164">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-1164">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-1165">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-1165">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-1166">1.0</span><span class="sxs-lookup"><span data-stu-id="c01fc-1166">1.0</span></span>|
|[<span data-ttu-id="c01fc-1167">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-1167">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-1168">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-1168">ReadItem</span></span>|
|[<span data-ttu-id="c01fc-1169">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-1169">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-1170">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-1170">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-1171">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-1171">Example</span></span>

<span data-ttu-id="c01fc-p174">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p174">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="c01fc-1175">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c01fc-1175">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="c01fc-1176">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1176">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="c01fc-p175">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p175">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c01fc-1181">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c01fc-1181">Parameters:</span></span>

|<span data-ttu-id="c01fc-1182">名前</span><span class="sxs-lookup"><span data-stu-id="c01fc-1182">Name</span></span>|<span data-ttu-id="c01fc-1183">型</span><span class="sxs-lookup"><span data-stu-id="c01fc-1183">Type</span></span>|<span data-ttu-id="c01fc-1184">属性</span><span class="sxs-lookup"><span data-stu-id="c01fc-1184">Attributes</span></span>|<span data-ttu-id="c01fc-1185">説明</span><span class="sxs-lookup"><span data-stu-id="c01fc-1185">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="c01fc-1186">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-1186">String</span></span>||<span data-ttu-id="c01fc-p176">削除する添付ファイルの識別子。文字列の最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p176">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="c01fc-1189">Object</span><span class="sxs-lookup"><span data-stu-id="c01fc-1189">Object</span></span>|<span data-ttu-id="c01fc-1190">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-1190">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-1191">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1191">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c01fc-1192">Object</span><span class="sxs-lookup"><span data-stu-id="c01fc-1192">Object</span></span>|<span data-ttu-id="c01fc-1193">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-1193">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-1194">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1194">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c01fc-1195">function</span><span class="sxs-lookup"><span data-stu-id="c01fc-1195">function</span></span>|<span data-ttu-id="c01fc-1196">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-1196">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-1197">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1197">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c01fc-1198">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1198">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c01fc-1199">エラー</span><span class="sxs-lookup"><span data-stu-id="c01fc-1199">Errors</span></span>

|<span data-ttu-id="c01fc-1200">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c01fc-1200">Error code</span></span>|<span data-ttu-id="c01fc-1201">説明</span><span class="sxs-lookup"><span data-stu-id="c01fc-1201">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="c01fc-1202">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1202">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c01fc-1203">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-1203">Requirements</span></span>

|<span data-ttu-id="c01fc-1204">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-1204">Requirement</span></span>|<span data-ttu-id="c01fc-1205">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-1205">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-1206">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-1206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-1207">1.1</span><span class="sxs-lookup"><span data-stu-id="c01fc-1207">1.1</span></span>|
|[<span data-ttu-id="c01fc-1208">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-1208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-1209">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-1209">ReadWriteItem</span></span>|
|[<span data-ttu-id="c01fc-1210">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-1210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-1211">Compose</span><span class="sxs-lookup"><span data-stu-id="c01fc-1211">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-1212">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-1212">Example</span></span>

<span data-ttu-id="c01fc-1213">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1213">The following code removes an attachment with an identifier of '0'.</span></span>

```js
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="c01fc-1214">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c01fc-1214">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="c01fc-1215">サポートされているイベントのイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1215">Removes an event handler for a</span></span>

<span data-ttu-id="c01fc-1216">現在、サポートされているイベントの種類は `Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged`、`Office.EventType.RecurrenceChanged` です。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1216">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="c01fc-1217">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c01fc-1217">Parameters:</span></span>

| <span data-ttu-id="c01fc-1218">名前</span><span class="sxs-lookup"><span data-stu-id="c01fc-1218">Name</span></span> | <span data-ttu-id="c01fc-1219">型</span><span class="sxs-lookup"><span data-stu-id="c01fc-1219">Type</span></span> | <span data-ttu-id="c01fc-1220">属性</span><span class="sxs-lookup"><span data-stu-id="c01fc-1220">Attributes</span></span> | <span data-ttu-id="c01fc-1221">説明</span><span class="sxs-lookup"><span data-stu-id="c01fc-1221">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c01fc-1222">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c01fc-1222">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c01fc-1223">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1223">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="c01fc-1224">Function</span><span class="sxs-lookup"><span data-stu-id="c01fc-1224">Function</span></span> || <span data-ttu-id="c01fc-p177">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`removeHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p177">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="c01fc-1228">Object</span><span class="sxs-lookup"><span data-stu-id="c01fc-1228">Object</span></span> | <span data-ttu-id="c01fc-1229">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-1229">&lt;optional&gt;</span></span> | <span data-ttu-id="c01fc-1230">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1230">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c01fc-1231">Object</span><span class="sxs-lookup"><span data-stu-id="c01fc-1231">Object</span></span> | <span data-ttu-id="c01fc-1232">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-1232">&lt;optional&gt;</span></span> | <span data-ttu-id="c01fc-1233">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1233">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c01fc-1234">function</span><span class="sxs-lookup"><span data-stu-id="c01fc-1234">function</span></span>| <span data-ttu-id="c01fc-1235">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-1235">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-1236">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1236">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c01fc-1237">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-1237">Requirements</span></span>

|<span data-ttu-id="c01fc-1238">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-1238">Requirement</span></span>| <span data-ttu-id="c01fc-1239">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-1239">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-1240">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-1240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c01fc-1241">1.7</span><span class="sxs-lookup"><span data-stu-id="c01fc-1241">-17</span></span> |
|[<span data-ttu-id="c01fc-1242">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-1242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c01fc-1243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-1243">ReadItem</span></span> |
|[<span data-ttu-id="c01fc-1244">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-1244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c01fc-1245">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="c01fc-1245">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="c01fc-1246">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-1246">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.removeHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="c01fc-1247">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="c01fc-1247">saveAsync([options], callback)</span></span>

<span data-ttu-id="c01fc-1248">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1248">Asynchronously saves an item.</span></span>

<span data-ttu-id="c01fc-p178">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p178">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="c01fc-1252">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1252">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="c01fc-1253">アイテムが同期されるまでに、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1253">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="c01fc-p180">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p180">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="c01fc-1257">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1257">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="c01fc-1258">Mac Outlook では、新規作成モードの会議で `saveAsync` をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1258">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="c01fc-1259">Mac Outlook では、会議で `saveAsync` を呼び出すとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1259">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="c01fc-1260">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状またはアップデートが常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1260">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c01fc-1261">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c01fc-1261">Parameters:</span></span>

|<span data-ttu-id="c01fc-1262">名前</span><span class="sxs-lookup"><span data-stu-id="c01fc-1262">Name</span></span>|<span data-ttu-id="c01fc-1263">型</span><span class="sxs-lookup"><span data-stu-id="c01fc-1263">Type</span></span>|<span data-ttu-id="c01fc-1264">属性</span><span class="sxs-lookup"><span data-stu-id="c01fc-1264">Attributes</span></span>|<span data-ttu-id="c01fc-1265">説明</span><span class="sxs-lookup"><span data-stu-id="c01fc-1265">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c01fc-1266">Object</span><span class="sxs-lookup"><span data-stu-id="c01fc-1266">Object</span></span>|<span data-ttu-id="c01fc-1267">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-1267">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-1268">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1268">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c01fc-1269">Object</span><span class="sxs-lookup"><span data-stu-id="c01fc-1269">Object</span></span>|<span data-ttu-id="c01fc-1270">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-1270">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-1271">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1271">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c01fc-1272">function</span><span class="sxs-lookup"><span data-stu-id="c01fc-1272">function</span></span>||<span data-ttu-id="c01fc-1273">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1273">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c01fc-1274">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1274">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c01fc-1275">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-1275">Requirements</span></span>

|<span data-ttu-id="c01fc-1276">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-1276">Requirement</span></span>|<span data-ttu-id="c01fc-1277">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-1277">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-1278">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-1278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-1279">1.3</span><span class="sxs-lookup"><span data-stu-id="c01fc-1279">1.3</span></span>|
|[<span data-ttu-id="c01fc-1280">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-1280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-1281">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-1281">ReadWriteItem</span></span>|
|[<span data-ttu-id="c01fc-1282">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-1282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-1283">Compose</span><span class="sxs-lookup"><span data-stu-id="c01fc-1283">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c01fc-1284">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-1284">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="c01fc-p182">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p182">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="c01fc-1287">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="c01fc-1287">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="c01fc-1288">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1288">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="c01fc-p183">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p183">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c01fc-1292">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c01fc-1292">Parameters:</span></span>

|<span data-ttu-id="c01fc-1293">名前</span><span class="sxs-lookup"><span data-stu-id="c01fc-1293">Name</span></span>|<span data-ttu-id="c01fc-1294">型</span><span class="sxs-lookup"><span data-stu-id="c01fc-1294">Type</span></span>|<span data-ttu-id="c01fc-1295">属性</span><span class="sxs-lookup"><span data-stu-id="c01fc-1295">Attributes</span></span>|<span data-ttu-id="c01fc-1296">説明</span><span class="sxs-lookup"><span data-stu-id="c01fc-1296">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="c01fc-1297">String</span><span class="sxs-lookup"><span data-stu-id="c01fc-1297">String</span></span>||<span data-ttu-id="c01fc-p184">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p184">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="c01fc-1301">Object</span><span class="sxs-lookup"><span data-stu-id="c01fc-1301">Object</span></span>|<span data-ttu-id="c01fc-1302">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-1303">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1303">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c01fc-1304">Object</span><span class="sxs-lookup"><span data-stu-id="c01fc-1304">Object</span></span>|<span data-ttu-id="c01fc-1305">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-1305">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-1306">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1306">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="c01fc-1307">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c01fc-1307">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="c01fc-1308">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c01fc-1308">&lt;optional&gt;</span></span>|<span data-ttu-id="c01fc-p185">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p185">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="c01fc-p186">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-p186">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="c01fc-1313">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1313">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="c01fc-1314">function</span><span class="sxs-lookup"><span data-stu-id="c01fc-1314">function</span></span>||<span data-ttu-id="c01fc-1315">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c01fc-1315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c01fc-1316">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-1316">Requirements</span></span>

|<span data-ttu-id="c01fc-1317">要件</span><span class="sxs-lookup"><span data-stu-id="c01fc-1317">Requirement</span></span>|<span data-ttu-id="c01fc-1318">値</span><span class="sxs-lookup"><span data-stu-id="c01fc-1318">Value</span></span>|
|---|---|
|[<span data-ttu-id="c01fc-1319">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c01fc-1319">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c01fc-1320">1.2</span><span class="sxs-lookup"><span data-stu-id="c01fc-1320">1.2</span></span>|
|[<span data-ttu-id="c01fc-1321">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c01fc-1321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c01fc-1322">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c01fc-1322">ReadWriteItem</span></span>|
|[<span data-ttu-id="c01fc-1323">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c01fc-1323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c01fc-1324">Compose</span><span class="sxs-lookup"><span data-stu-id="c01fc-1324">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c01fc-1325">例</span><span class="sxs-lookup"><span data-stu-id="c01fc-1325">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```