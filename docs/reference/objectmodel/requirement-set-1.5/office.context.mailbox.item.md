
# <a name="item"></a><span data-ttu-id="8c151-101">item</span><span class="sxs-lookup"><span data-stu-id="8c151-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="8c151-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="8c151-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="8c151-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="8c151-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c151-105">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-105">Requirements</span></span>

|<span data-ttu-id="8c151-106">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-106">Requirement</span></span>| <span data-ttu-id="8c151-107">値</span><span class="sxs-lookup"><span data-stu-id="8c151-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-109">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-109">1.0</span></span>|
|[<span data-ttu-id="8c151-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="8c151-111">Restricted</span></span>|
|[<span data-ttu-id="8c151-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-113">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8c151-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="8c151-114">Members and methods</span></span>

| <span data-ttu-id="8c151-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-115">Member</span></span> | <span data-ttu-id="8c151-116">種類</span><span class="sxs-lookup"><span data-stu-id="8c151-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8c151-117">attachments</span><span class="sxs-lookup"><span data-stu-id="8c151-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails) | <span data-ttu-id="8c151-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-118">Member</span></span> |
| [<span data-ttu-id="8c151-119">bcc</span><span class="sxs-lookup"><span data-stu-id="8c151-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="8c151-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-120">Member</span></span> |
| [<span data-ttu-id="8c151-121">body</span><span class="sxs-lookup"><span data-stu-id="8c151-121">body</span></span>](#body-bodyjavascriptapioutlook15officebody) | <span data-ttu-id="8c151-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-122">Member</span></span> |
| [<span data-ttu-id="8c151-123">cc</span><span class="sxs-lookup"><span data-stu-id="8c151-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="8c151-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-124">Member</span></span> |
| [<span data-ttu-id="8c151-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="8c151-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="8c151-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-126">Member</span></span> |
| [<span data-ttu-id="8c151-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="8c151-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="8c151-128">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-128">Member</span></span> |
| [<span data-ttu-id="8c151-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="8c151-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="8c151-130">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-130">Member</span></span> |
| [<span data-ttu-id="8c151-131">end</span><span class="sxs-lookup"><span data-stu-id="8c151-131">end</span></span>](#end-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="8c151-132">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-132">Member</span></span> |
| [<span data-ttu-id="8c151-133">from</span><span class="sxs-lookup"><span data-stu-id="8c151-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="8c151-134">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-134">Member</span></span> |
| [<span data-ttu-id="8c151-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="8c151-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="8c151-136">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-136">Member</span></span> |
| [<span data-ttu-id="8c151-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="8c151-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="8c151-138">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-138">Member</span></span> |
| [<span data-ttu-id="8c151-139">itemId</span><span class="sxs-lookup"><span data-stu-id="8c151-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="8c151-140">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-140">Member</span></span> |
| [<span data-ttu-id="8c151-141">itemType</span><span class="sxs-lookup"><span data-stu-id="8c151-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) | <span data-ttu-id="8c151-142">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-142">Member</span></span> |
| [<span data-ttu-id="8c151-143">location</span><span class="sxs-lookup"><span data-stu-id="8c151-143">location</span></span>](#location-stringlocationjavascriptapioutlook15officelocation) | <span data-ttu-id="8c151-144">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-144">Member</span></span> |
| [<span data-ttu-id="8c151-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="8c151-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="8c151-146">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-146">Member</span></span> |
| [<span data-ttu-id="8c151-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="8c151-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages) | <span data-ttu-id="8c151-148">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-148">Member</span></span> |
| [<span data-ttu-id="8c151-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="8c151-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="8c151-150">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-150">Member</span></span> |
| [<span data-ttu-id="8c151-151">organizer</span><span class="sxs-lookup"><span data-stu-id="8c151-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="8c151-152">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-152">Member</span></span> |
| [<span data-ttu-id="8c151-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="8c151-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="8c151-154">Member</span><span class="sxs-lookup"><span data-stu-id="8c151-154">Member</span></span> |
| [<span data-ttu-id="8c151-155">sender</span><span class="sxs-lookup"><span data-stu-id="8c151-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="8c151-156">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-156">Member</span></span> |
| [<span data-ttu-id="8c151-157">start</span><span class="sxs-lookup"><span data-stu-id="8c151-157">start</span></span>](#start-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="8c151-158">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-158">Member</span></span> |
| [<span data-ttu-id="8c151-159">subject</span><span class="sxs-lookup"><span data-stu-id="8c151-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook15officesubject) | <span data-ttu-id="8c151-160">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-160">Member</span></span> |
| [<span data-ttu-id="8c151-161">to</span><span class="sxs-lookup"><span data-stu-id="8c151-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="8c151-162">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-162">Member</span></span> |
| [<span data-ttu-id="8c151-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8c151-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="8c151-164">メソッド</span><span class="sxs-lookup"><span data-stu-id="8c151-164">Method</span></span> |
| [<span data-ttu-id="8c151-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8c151-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="8c151-166">メソッド</span><span class="sxs-lookup"><span data-stu-id="8c151-166">Method</span></span> |
| [<span data-ttu-id="8c151-167">close</span><span class="sxs-lookup"><span data-stu-id="8c151-167">close</span></span>](#close) | <span data-ttu-id="8c151-168">メソッド</span><span class="sxs-lookup"><span data-stu-id="8c151-168">Method</span></span> |
| [<span data-ttu-id="8c151-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="8c151-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="8c151-170">メソッド</span><span class="sxs-lookup"><span data-stu-id="8c151-170">Method</span></span> |
| [<span data-ttu-id="8c151-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="8c151-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="8c151-172">メソッド</span><span class="sxs-lookup"><span data-stu-id="8c151-172">Method</span></span> |
| [<span data-ttu-id="8c151-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="8c151-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook15officeentities) | <span data-ttu-id="8c151-174">メソッド</span><span class="sxs-lookup"><span data-stu-id="8c151-174">Method</span></span> |
| [<span data-ttu-id="8c151-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="8c151-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="8c151-176">メソッド</span><span class="sxs-lookup"><span data-stu-id="8c151-176">Method</span></span> |
| [<span data-ttu-id="8c151-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="8c151-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="8c151-178">メソッド</span><span class="sxs-lookup"><span data-stu-id="8c151-178">Method</span></span> |
| [<span data-ttu-id="8c151-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="8c151-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="8c151-180">メソッド</span><span class="sxs-lookup"><span data-stu-id="8c151-180">Method</span></span> |
| [<span data-ttu-id="8c151-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="8c151-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="8c151-182">メソッド</span><span class="sxs-lookup"><span data-stu-id="8c151-182">Method</span></span> |
| [<span data-ttu-id="8c151-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="8c151-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="8c151-184">メソッド</span><span class="sxs-lookup"><span data-stu-id="8c151-184">Method</span></span> |
| [<span data-ttu-id="8c151-185">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="8c151-185">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="8c151-186">メソッド</span><span class="sxs-lookup"><span data-stu-id="8c151-186">Method</span></span> |
| [<span data-ttu-id="8c151-187">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8c151-187">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="8c151-188">メソッド</span><span class="sxs-lookup"><span data-stu-id="8c151-188">Method</span></span> |
| [<span data-ttu-id="8c151-189">saveAsync</span><span class="sxs-lookup"><span data-stu-id="8c151-189">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="8c151-190">メソッド</span><span class="sxs-lookup"><span data-stu-id="8c151-190">Method</span></span> |
| [<span data-ttu-id="8c151-191">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="8c151-191">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="8c151-192">メソッド</span><span class="sxs-lookup"><span data-stu-id="8c151-192">Method</span></span> |

### <a name="example"></a><span data-ttu-id="8c151-193">例</span><span class="sxs-lookup"><span data-stu-id="8c151-193">Example</span></span>

<span data-ttu-id="8c151-194">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8c151-194">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="8c151-195">メンバー</span><span class="sxs-lookup"><span data-stu-id="8c151-195">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="8c151-196">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="8c151-196">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="8c151-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8c151-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8c151-199">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="8c151-199">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="8c151-200">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8c151-200">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-201">種類:</span><span class="sxs-lookup"><span data-stu-id="8c151-201">Type:</span></span>

*   <span data-ttu-id="8c151-202">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="8c151-202">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="8c151-203">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-203">Requirements</span></span>

|<span data-ttu-id="8c151-204">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-204">Requirement</span></span>| <span data-ttu-id="8c151-205">値</span><span class="sxs-lookup"><span data-stu-id="8c151-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-206">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-207">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-207">1.0</span></span>|
|[<span data-ttu-id="8c151-208">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-209">ReadItem</span></span>|
|[<span data-ttu-id="8c151-210">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-211">読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-211">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-212">例</span><span class="sxs-lookup"><span data-stu-id="8c151-212">Example</span></span>

<span data-ttu-id="8c151-213">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="8c151-213">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="8c151-214">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8c151-214">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="8c151-215">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="8c151-215">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="8c151-216">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8c151-216">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-217">型:</span><span class="sxs-lookup"><span data-stu-id="8c151-217">Type:</span></span>

*   [<span data-ttu-id="8c151-218">Recipients</span><span class="sxs-lookup"><span data-stu-id="8c151-218">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="8c151-219">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-219">Requirements</span></span>

|<span data-ttu-id="8c151-220">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-220">Requirement</span></span>| <span data-ttu-id="8c151-221">値</span><span class="sxs-lookup"><span data-stu-id="8c151-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-222">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-222">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-223">1.1</span><span class="sxs-lookup"><span data-stu-id="8c151-223">1.1</span></span>|
|[<span data-ttu-id="8c151-224">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-224">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-225">ReadItem</span></span>|
|[<span data-ttu-id="8c151-226">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-226">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-227">作成</span><span class="sxs-lookup"><span data-stu-id="8c151-227">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-228">例</span><span class="sxs-lookup"><span data-stu-id="8c151-228">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="8c151-229">body :[Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="8c151-229">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="8c151-230">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="8c151-230">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-231">型:</span><span class="sxs-lookup"><span data-stu-id="8c151-231">Type:</span></span>

*   [<span data-ttu-id="8c151-232">Body</span><span class="sxs-lookup"><span data-stu-id="8c151-232">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="8c151-233">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-233">Requirements</span></span>

|<span data-ttu-id="8c151-234">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-234">Requirement</span></span>| <span data-ttu-id="8c151-235">値</span><span class="sxs-lookup"><span data-stu-id="8c151-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-236">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-237">1.1</span><span class="sxs-lookup"><span data-stu-id="8c151-237">1.1</span></span>|
|[<span data-ttu-id="8c151-238">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-238">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-239">ReadItem</span></span>|
|[<span data-ttu-id="8c151-240">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-240">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-241">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-241">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="8c151-242">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8c151-242">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="8c151-243">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8c151-243">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="8c151-244">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="8c151-244">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8c151-245">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8c151-245">Read mode</span></span>

<span data-ttu-id="8c151-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8c151-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8c151-248">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8c151-248">Compose mode</span></span>

<span data-ttu-id="8c151-249">`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-249">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-250">種類:</span><span class="sxs-lookup"><span data-stu-id="8c151-250">Type:</span></span>

*   <span data-ttu-id="8c151-251">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8c151-251">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c151-252">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-252">Requirements</span></span>

|<span data-ttu-id="8c151-253">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-253">Requirement</span></span>| <span data-ttu-id="8c151-254">値</span><span class="sxs-lookup"><span data-stu-id="8c151-254">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-255">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-255">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-256">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-256">1.0</span></span>|
|[<span data-ttu-id="8c151-257">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-257">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-258">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-258">ReadItem</span></span>|
|[<span data-ttu-id="8c151-259">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-259">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-260">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-260">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-261">例</span><span class="sxs-lookup"><span data-stu-id="8c151-261">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="8c151-262">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="8c151-262">(nullable) conversationId :String</span></span>

<span data-ttu-id="8c151-263">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="8c151-263">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="8c151-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="8c151-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="8c151-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-268">型:</span><span class="sxs-lookup"><span data-stu-id="8c151-268">Type:</span></span>

*   <span data-ttu-id="8c151-269">String</span><span class="sxs-lookup"><span data-stu-id="8c151-269">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c151-270">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-270">Requirements</span></span>

|<span data-ttu-id="8c151-271">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-271">Requirement</span></span>| <span data-ttu-id="8c151-272">値</span><span class="sxs-lookup"><span data-stu-id="8c151-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-273">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-274">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-274">1.0</span></span>|
|[<span data-ttu-id="8c151-275">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-276">ReadItem</span></span>|
|[<span data-ttu-id="8c151-277">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-278">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-278">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="8c151-279">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="8c151-279">dateTimeCreated :Date</span></span>

<span data-ttu-id="8c151-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8c151-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-282">型:</span><span class="sxs-lookup"><span data-stu-id="8c151-282">Type:</span></span>

*   <span data-ttu-id="8c151-283">日付</span><span class="sxs-lookup"><span data-stu-id="8c151-283">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c151-284">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-284">Requirements</span></span>

|<span data-ttu-id="8c151-285">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-285">Requirement</span></span>| <span data-ttu-id="8c151-286">値</span><span class="sxs-lookup"><span data-stu-id="8c151-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-287">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-287">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-288">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-288">1.0</span></span>|
|[<span data-ttu-id="8c151-289">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-289">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-290">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-290">ReadItem</span></span>|
|[<span data-ttu-id="8c151-291">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-291">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-292">読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-292">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-293">例</span><span class="sxs-lookup"><span data-stu-id="8c151-293">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="8c151-294">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="8c151-294">dateTimeModified :Date</span></span>

<span data-ttu-id="8c151-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8c151-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8c151-297">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8c151-297">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-298">型:</span><span class="sxs-lookup"><span data-stu-id="8c151-298">Type:</span></span>

*   <span data-ttu-id="8c151-299">日付</span><span class="sxs-lookup"><span data-stu-id="8c151-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c151-300">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-300">Requirements</span></span>

|<span data-ttu-id="8c151-301">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-301">Requirement</span></span>| <span data-ttu-id="8c151-302">値</span><span class="sxs-lookup"><span data-stu-id="8c151-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-303">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-304">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-304">1.0</span></span>|
|[<span data-ttu-id="8c151-305">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-306">ReadItem</span></span>|
|[<span data-ttu-id="8c151-307">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-308">読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-309">例</span><span class="sxs-lookup"><span data-stu-id="8c151-309">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="8c151-310">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="8c151-310">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="8c151-311">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8c151-311">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="8c151-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="8c151-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8c151-314">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8c151-314">Read mode</span></span>

<span data-ttu-id="8c151-315">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-315">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8c151-316">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8c151-316">Compose mode</span></span>

<span data-ttu-id="8c151-317">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-317">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="8c151-318">[`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8c151-318">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-319">型:</span><span class="sxs-lookup"><span data-stu-id="8c151-319">Type:</span></span>

*   <span data-ttu-id="8c151-320">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="8c151-320">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c151-321">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-321">Requirements</span></span>

|<span data-ttu-id="8c151-322">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-322">Requirement</span></span>| <span data-ttu-id="8c151-323">値</span><span class="sxs-lookup"><span data-stu-id="8c151-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-324">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-325">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-325">1.0</span></span>|
|[<span data-ttu-id="8c151-326">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-327">ReadItem</span></span>|
|[<span data-ttu-id="8c151-328">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-329">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-329">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-330">例</span><span class="sxs-lookup"><span data-stu-id="8c151-330">Example</span></span>

<span data-ttu-id="8c151-331">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="8c151-331">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="8c151-332">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="8c151-332">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="8c151-p112">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8c151-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="8c151-p113">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="8c151-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8c151-337">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="8c151-337">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-338">種類:</span><span class="sxs-lookup"><span data-stu-id="8c151-338">Type:</span></span>

*   [<span data-ttu-id="8c151-339">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8c151-339">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="8c151-340">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-340">Requirements</span></span>

|<span data-ttu-id="8c151-341">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-341">Requirement</span></span>| <span data-ttu-id="8c151-342">値</span><span class="sxs-lookup"><span data-stu-id="8c151-342">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-343">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-343">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-344">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-344">1.0</span></span>|
|[<span data-ttu-id="8c151-345">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-345">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-346">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-346">ReadItem</span></span>|
|[<span data-ttu-id="8c151-347">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-347">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-348">読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-348">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="8c151-349">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="8c151-349">internetMessageId :String</span></span>

<span data-ttu-id="8c151-p114">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8c151-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-352">型:</span><span class="sxs-lookup"><span data-stu-id="8c151-352">Type:</span></span>

*   <span data-ttu-id="8c151-353">String</span><span class="sxs-lookup"><span data-stu-id="8c151-353">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c151-354">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-354">Requirements</span></span>

|<span data-ttu-id="8c151-355">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-355">Requirement</span></span>| <span data-ttu-id="8c151-356">値</span><span class="sxs-lookup"><span data-stu-id="8c151-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-357">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-358">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-358">1.0</span></span>|
|[<span data-ttu-id="8c151-359">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-359">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-360">ReadItem</span></span>|
|[<span data-ttu-id="8c151-361">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-361">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-362">読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-362">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-363">例</span><span class="sxs-lookup"><span data-stu-id="8c151-363">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="8c151-364">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="8c151-364">itemClass :String</span></span>

<span data-ttu-id="8c151-p115">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8c151-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="8c151-p116">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="8c151-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="8c151-369">型</span><span class="sxs-lookup"><span data-stu-id="8c151-369">Type</span></span> | <span data-ttu-id="8c151-370">説明</span><span class="sxs-lookup"><span data-stu-id="8c151-370">Description</span></span> | <span data-ttu-id="8c151-371">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="8c151-371">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="8c151-372">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="8c151-372">Appointment items</span></span> | <span data-ttu-id="8c151-373">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8c151-373">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="8c151-374">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="8c151-374">Message items</span></span> | <span data-ttu-id="8c151-375">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="8c151-375">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="8c151-376">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="8c151-376">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-377">型:</span><span class="sxs-lookup"><span data-stu-id="8c151-377">Type:</span></span>

*   <span data-ttu-id="8c151-378">String</span><span class="sxs-lookup"><span data-stu-id="8c151-378">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c151-379">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-379">Requirements</span></span>

|<span data-ttu-id="8c151-380">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-380">Requirement</span></span>| <span data-ttu-id="8c151-381">値</span><span class="sxs-lookup"><span data-stu-id="8c151-381">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-382">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-382">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-383">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-383">1.0</span></span>|
|[<span data-ttu-id="8c151-384">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-384">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-385">ReadItem</span></span>|
|[<span data-ttu-id="8c151-386">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-386">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-387">読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-387">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-388">例</span><span class="sxs-lookup"><span data-stu-id="8c151-388">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="8c151-389">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="8c151-389">(nullable) itemId :String</span></span>

<span data-ttu-id="8c151-p117">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8c151-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8c151-392">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="8c151-392">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="8c151-393">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="8c151-393">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="8c151-394">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8c151-394">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="8c151-395">詳細は、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8c151-395">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="8c151-p119">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-398">種類:</span><span class="sxs-lookup"><span data-stu-id="8c151-398">Type:</span></span>

*   <span data-ttu-id="8c151-399">String</span><span class="sxs-lookup"><span data-stu-id="8c151-399">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c151-400">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-400">Requirements</span></span>

|<span data-ttu-id="8c151-401">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-401">Requirement</span></span>| <span data-ttu-id="8c151-402">値</span><span class="sxs-lookup"><span data-stu-id="8c151-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-403">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-403">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-404">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-404">1.0</span></span>|
|[<span data-ttu-id="8c151-405">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-405">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-406">ReadItem</span></span>|
|[<span data-ttu-id="8c151-407">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-407">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-408">読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-408">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-409">例</span><span class="sxs-lookup"><span data-stu-id="8c151-409">Example</span></span>

<span data-ttu-id="8c151-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="8c151-412">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="8c151-412">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="8c151-413">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="8c151-413">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="8c151-414">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="8c151-414">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-415">型:</span><span class="sxs-lookup"><span data-stu-id="8c151-415">Type:</span></span>

*   [<span data-ttu-id="8c151-416">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="8c151-416">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="8c151-417">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-417">Requirements</span></span>

|<span data-ttu-id="8c151-418">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-418">Requirement</span></span>| <span data-ttu-id="8c151-419">値</span><span class="sxs-lookup"><span data-stu-id="8c151-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-420">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-420">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-421">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-421">1.0</span></span>|
|[<span data-ttu-id="8c151-422">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-422">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-423">ReadItem</span></span>|
|[<span data-ttu-id="8c151-424">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-424">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-425">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-425">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-426">例</span><span class="sxs-lookup"><span data-stu-id="8c151-426">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="8c151-427">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="8c151-427">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="8c151-428">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8c151-428">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8c151-429">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8c151-429">Read mode</span></span>

<span data-ttu-id="8c151-430">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-430">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8c151-431">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8c151-431">Compose mode</span></span>

<span data-ttu-id="8c151-432">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-432">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-433">型:</span><span class="sxs-lookup"><span data-stu-id="8c151-433">Type:</span></span>

*   <span data-ttu-id="8c151-434">String | [Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="8c151-434">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c151-435">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-435">Requirements</span></span>

|<span data-ttu-id="8c151-436">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-436">Requirement</span></span>| <span data-ttu-id="8c151-437">値</span><span class="sxs-lookup"><span data-stu-id="8c151-437">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-438">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-438">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-439">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-439">1.0</span></span>|
|[<span data-ttu-id="8c151-440">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-440">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-441">ReadItem</span></span>|
|[<span data-ttu-id="8c151-442">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-442">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-443">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-443">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-444">例</span><span class="sxs-lookup"><span data-stu-id="8c151-444">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="8c151-445">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="8c151-445">normalizedSubject :String</span></span>

<span data-ttu-id="8c151-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8c151-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="8c151-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="8c151-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-450">型:</span><span class="sxs-lookup"><span data-stu-id="8c151-450">Type:</span></span>

*   <span data-ttu-id="8c151-451">String</span><span class="sxs-lookup"><span data-stu-id="8c151-451">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c151-452">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-452">Requirements</span></span>

|<span data-ttu-id="8c151-453">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-453">Requirement</span></span>| <span data-ttu-id="8c151-454">値</span><span class="sxs-lookup"><span data-stu-id="8c151-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-455">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-455">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-456">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-456">1.0</span></span>|
|[<span data-ttu-id="8c151-457">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-457">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-458">ReadItem</span></span>|
|[<span data-ttu-id="8c151-459">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-459">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-460">読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-460">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-461">例</span><span class="sxs-lookup"><span data-stu-id="8c151-461">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="8c151-462">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="8c151-462">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="8c151-463">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="8c151-463">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-464">型:</span><span class="sxs-lookup"><span data-stu-id="8c151-464">Type:</span></span>

*   [<span data-ttu-id="8c151-465">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="8c151-465">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="8c151-466">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-466">Requirements</span></span>

|<span data-ttu-id="8c151-467">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-467">Requirement</span></span>| <span data-ttu-id="8c151-468">値</span><span class="sxs-lookup"><span data-stu-id="8c151-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-469">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-470">1.3</span><span class="sxs-lookup"><span data-stu-id="8c151-470">1.3</span></span>|
|[<span data-ttu-id="8c151-471">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-471">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-472">ReadItem</span></span>|
|[<span data-ttu-id="8c151-473">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-473">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-474">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-474">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="8c151-475">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8c151-475">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="8c151-476">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8c151-476">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="8c151-477">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="8c151-477">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8c151-478">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8c151-478">Read mode</span></span>

<span data-ttu-id="8c151-479">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-479">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8c151-480">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8c151-480">Compose mode</span></span>

<span data-ttu-id="8c151-481">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-481">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-482">種類:</span><span class="sxs-lookup"><span data-stu-id="8c151-482">Type:</span></span>

*   <span data-ttu-id="8c151-483">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8c151-483">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c151-484">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-484">Requirements</span></span>

|<span data-ttu-id="8c151-485">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-485">Requirement</span></span>| <span data-ttu-id="8c151-486">値</span><span class="sxs-lookup"><span data-stu-id="8c151-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-487">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-487">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-488">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-488">1.0</span></span>|
|[<span data-ttu-id="8c151-489">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-489">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-490">ReadItem</span></span>|
|[<span data-ttu-id="8c151-491">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-491">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-492">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-492">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-493">例</span><span class="sxs-lookup"><span data-stu-id="8c151-493">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="8c151-494">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="8c151-494">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="8c151-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8c151-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-497">型:</span><span class="sxs-lookup"><span data-stu-id="8c151-497">Type:</span></span>

*   [<span data-ttu-id="8c151-498">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8c151-498">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="8c151-499">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-499">Requirements</span></span>

|<span data-ttu-id="8c151-500">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-500">Requirement</span></span>| <span data-ttu-id="8c151-501">値</span><span class="sxs-lookup"><span data-stu-id="8c151-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-502">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-503">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-503">1.0</span></span>|
|[<span data-ttu-id="8c151-504">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-505">ReadItem</span></span>|
|[<span data-ttu-id="8c151-506">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-507">読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-507">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-508">例</span><span class="sxs-lookup"><span data-stu-id="8c151-508">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="8c151-509">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8c151-509">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="8c151-510">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8c151-510">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="8c151-511">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="8c151-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8c151-512">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8c151-512">Read mode</span></span>

<span data-ttu-id="8c151-513">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-513">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8c151-514">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8c151-514">Compose mode</span></span>

<span data-ttu-id="8c151-515">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-515">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-516">種類:</span><span class="sxs-lookup"><span data-stu-id="8c151-516">Type:</span></span>

*   <span data-ttu-id="8c151-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8c151-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c151-518">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-518">Requirements</span></span>

|<span data-ttu-id="8c151-519">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-519">Requirement</span></span>| <span data-ttu-id="8c151-520">値</span><span class="sxs-lookup"><span data-stu-id="8c151-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-521">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-522">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-522">1.0</span></span>|
|[<span data-ttu-id="8c151-523">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-524">ReadItem</span></span>|
|[<span data-ttu-id="8c151-525">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-526">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-527">例</span><span class="sxs-lookup"><span data-stu-id="8c151-527">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="8c151-528">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="8c151-528">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="8c151-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8c151-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="8c151-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="8c151-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8c151-533">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="8c151-533">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-534">種類:</span><span class="sxs-lookup"><span data-stu-id="8c151-534">Type:</span></span>

*   [<span data-ttu-id="8c151-535">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8c151-535">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="8c151-536">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-536">Requirements</span></span>

|<span data-ttu-id="8c151-537">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-537">Requirement</span></span>| <span data-ttu-id="8c151-538">値</span><span class="sxs-lookup"><span data-stu-id="8c151-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-539">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-540">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-540">1.0</span></span>|
|[<span data-ttu-id="8c151-541">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-541">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-542">ReadItem</span></span>|
|[<span data-ttu-id="8c151-543">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-543">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-544">読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-544">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-545">例</span><span class="sxs-lookup"><span data-stu-id="8c151-545">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="8c151-546">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="8c151-546">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="8c151-547">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8c151-547">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="8c151-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="8c151-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8c151-550">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8c151-550">Read mode</span></span>

<span data-ttu-id="8c151-551">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-551">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8c151-552">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8c151-552">Compose mode</span></span>

<span data-ttu-id="8c151-553">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-553">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="8c151-554">[`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8c151-554">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-555">型:</span><span class="sxs-lookup"><span data-stu-id="8c151-555">Type:</span></span>

*   <span data-ttu-id="8c151-556">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="8c151-556">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c151-557">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-557">Requirements</span></span>

|<span data-ttu-id="8c151-558">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-558">Requirement</span></span>| <span data-ttu-id="8c151-559">値</span><span class="sxs-lookup"><span data-stu-id="8c151-559">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-560">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-560">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-561">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-561">1.0</span></span>|
|[<span data-ttu-id="8c151-562">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-562">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-563">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-563">ReadItem</span></span>|
|[<span data-ttu-id="8c151-564">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-564">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-565">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-565">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-566">例</span><span class="sxs-lookup"><span data-stu-id="8c151-566">Example</span></span>

<span data-ttu-id="8c151-567">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="8c151-567">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="8c151-568">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="8c151-568">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="8c151-569">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8c151-569">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="8c151-570">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8c151-570">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8c151-571">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8c151-571">Read mode</span></span>

<span data-ttu-id="8c151-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="8c151-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="8c151-574">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8c151-574">Compose mode</span></span>

<span data-ttu-id="8c151-575">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-575">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8c151-576">型:</span><span class="sxs-lookup"><span data-stu-id="8c151-576">Type:</span></span>

*   <span data-ttu-id="8c151-577">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="8c151-577">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c151-578">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-578">Requirements</span></span>

|<span data-ttu-id="8c151-579">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-579">Requirement</span></span>| <span data-ttu-id="8c151-580">値</span><span class="sxs-lookup"><span data-stu-id="8c151-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-581">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-581">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-582">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-582">1.0</span></span>|
|[<span data-ttu-id="8c151-583">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-583">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-584">ReadItem</span></span>|
|[<span data-ttu-id="8c151-585">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-585">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-586">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-586">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="8c151-587">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8c151-587">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="8c151-588">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8c151-588">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="8c151-589">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="8c151-589">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8c151-590">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8c151-590">Read mode</span></span>

<span data-ttu-id="8c151-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8c151-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8c151-593">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8c151-593">Compose mode</span></span>

<span data-ttu-id="8c151-594">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-594">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="8c151-595">種類:</span><span class="sxs-lookup"><span data-stu-id="8c151-595">Type:</span></span>

*   <span data-ttu-id="8c151-596">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8c151-596">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c151-597">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-597">Requirements</span></span>

|<span data-ttu-id="8c151-598">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-598">Requirement</span></span>| <span data-ttu-id="8c151-599">値</span><span class="sxs-lookup"><span data-stu-id="8c151-599">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-600">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-600">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-601">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-601">1.0</span></span>|
|[<span data-ttu-id="8c151-602">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-602">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-603">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-603">ReadItem</span></span>|
|[<span data-ttu-id="8c151-604">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-604">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-605">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-605">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-606">例</span><span class="sxs-lookup"><span data-stu-id="8c151-606">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="8c151-607">メソッド</span><span class="sxs-lookup"><span data-stu-id="8c151-607">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="8c151-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8c151-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8c151-609">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="8c151-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="8c151-610">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="8c151-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="8c151-611">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="8c151-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c151-612">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8c151-612">Parameters:</span></span>

|<span data-ttu-id="8c151-613">名前</span><span class="sxs-lookup"><span data-stu-id="8c151-613">Name</span></span>| <span data-ttu-id="8c151-614">型</span><span class="sxs-lookup"><span data-stu-id="8c151-614">Type</span></span>| <span data-ttu-id="8c151-615">属性</span><span class="sxs-lookup"><span data-stu-id="8c151-615">Attributes</span></span>| <span data-ttu-id="8c151-616">説明</span><span class="sxs-lookup"><span data-stu-id="8c151-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="8c151-617">String</span><span class="sxs-lookup"><span data-stu-id="8c151-617">String</span></span>||<span data-ttu-id="8c151-p132">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="8c151-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="8c151-620">String</span><span class="sxs-lookup"><span data-stu-id="8c151-620">String</span></span>||<span data-ttu-id="8c151-p133">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="8c151-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="8c151-623">Object</span><span class="sxs-lookup"><span data-stu-id="8c151-623">Object</span></span>| <span data-ttu-id="8c151-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-624">&lt;optional&gt;</span></span>|<span data-ttu-id="8c151-625">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8c151-625">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="8c151-626">Object</span><span class="sxs-lookup"><span data-stu-id="8c151-626">Object</span></span> | <span data-ttu-id="8c151-627">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-627">&lt;optional&gt;</span></span> | <span data-ttu-id="8c151-628">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8c151-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="8c151-629">Boolean</span><span class="sxs-lookup"><span data-stu-id="8c151-629">Boolean</span></span> | <span data-ttu-id="8c151-630">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-630">&lt;optional&gt;</span></span> | <span data-ttu-id="8c151-631">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="8c151-631">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="8c151-632">function</span><span class="sxs-lookup"><span data-stu-id="8c151-632">function</span></span>| <span data-ttu-id="8c151-633">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-633">&lt;optional&gt;</span></span>|<span data-ttu-id="8c151-634">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8c151-635">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-635">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8c151-636">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="8c151-636">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8c151-637">エラー</span><span class="sxs-lookup"><span data-stu-id="8c151-637">Errors</span></span>

| <span data-ttu-id="8c151-638">エラー コード</span><span class="sxs-lookup"><span data-stu-id="8c151-638">Error code</span></span> | <span data-ttu-id="8c151-639">説明</span><span class="sxs-lookup"><span data-stu-id="8c151-639">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="8c151-640">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="8c151-640">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="8c151-641">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="8c151-641">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="8c151-642">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="8c151-642">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8c151-643">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-643">Requirements</span></span>

|<span data-ttu-id="8c151-644">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-644">Requirement</span></span>| <span data-ttu-id="8c151-645">値</span><span class="sxs-lookup"><span data-stu-id="8c151-645">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-646">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-646">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-647">1.1</span><span class="sxs-lookup"><span data-stu-id="8c151-647">1.1</span></span>|
|[<span data-ttu-id="8c151-648">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-648">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-649">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8c151-649">ReadWriteItem</span></span>|
|[<span data-ttu-id="8c151-650">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-650">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-651">作成</span><span class="sxs-lookup"><span data-stu-id="8c151-651">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8c151-652">例</span><span class="sxs-lookup"><span data-stu-id="8c151-652">Examples</span></span>

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

<span data-ttu-id="8c151-653">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="8c151-653">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="8c151-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8c151-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8c151-655">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="8c151-655">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="8c151-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="8c151-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="8c151-659">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="8c151-659">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="8c151-660">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="8c151-660">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c151-661">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8c151-661">Parameters:</span></span>

|<span data-ttu-id="8c151-662">名前</span><span class="sxs-lookup"><span data-stu-id="8c151-662">Name</span></span>| <span data-ttu-id="8c151-663">型</span><span class="sxs-lookup"><span data-stu-id="8c151-663">Type</span></span>| <span data-ttu-id="8c151-664">属性</span><span class="sxs-lookup"><span data-stu-id="8c151-664">Attributes</span></span>| <span data-ttu-id="8c151-665">説明</span><span class="sxs-lookup"><span data-stu-id="8c151-665">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="8c151-666">String</span><span class="sxs-lookup"><span data-stu-id="8c151-666">String</span></span>||<span data-ttu-id="8c151-p135">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="8c151-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="8c151-669">String</span><span class="sxs-lookup"><span data-stu-id="8c151-669">String</span></span>||<span data-ttu-id="8c151-p136">添付するアイテムの件名。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="8c151-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="8c151-672">Object</span><span class="sxs-lookup"><span data-stu-id="8c151-672">Object</span></span>| <span data-ttu-id="8c151-673">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-673">&lt;optional&gt;</span></span>|<span data-ttu-id="8c151-674">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8c151-674">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8c151-675">Object</span><span class="sxs-lookup"><span data-stu-id="8c151-675">Object</span></span>| <span data-ttu-id="8c151-676">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-676">&lt;optional&gt;</span></span>|<span data-ttu-id="8c151-677">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8c151-677">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8c151-678">function</span><span class="sxs-lookup"><span data-stu-id="8c151-678">function</span></span>| <span data-ttu-id="8c151-679">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-679">&lt;optional&gt;</span></span>|<span data-ttu-id="8c151-680">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-680">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8c151-681">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-681">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8c151-682">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="8c151-682">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8c151-683">エラー</span><span class="sxs-lookup"><span data-stu-id="8c151-683">Errors</span></span>

| <span data-ttu-id="8c151-684">エラー コード</span><span class="sxs-lookup"><span data-stu-id="8c151-684">Error code</span></span> | <span data-ttu-id="8c151-685">説明</span><span class="sxs-lookup"><span data-stu-id="8c151-685">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="8c151-686">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="8c151-686">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8c151-687">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-687">Requirements</span></span>

|<span data-ttu-id="8c151-688">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-688">Requirement</span></span>| <span data-ttu-id="8c151-689">値</span><span class="sxs-lookup"><span data-stu-id="8c151-689">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-690">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-690">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-691">1.1</span><span class="sxs-lookup"><span data-stu-id="8c151-691">1.1</span></span>|
|[<span data-ttu-id="8c151-692">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-692">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-693">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8c151-693">ReadWriteItem</span></span>|
|[<span data-ttu-id="8c151-694">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-694">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-695">作成</span><span class="sxs-lookup"><span data-stu-id="8c151-695">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-696">例</span><span class="sxs-lookup"><span data-stu-id="8c151-696">Example</span></span>

<span data-ttu-id="8c151-697">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-697">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="8c151-698">close()</span><span class="sxs-lookup"><span data-stu-id="8c151-698">close()</span></span>

<span data-ttu-id="8c151-699">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="8c151-699">Closes the current item that is being composed.</span></span>

<span data-ttu-id="8c151-p137">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="8c151-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="8c151-702">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-702">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="8c151-703">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="8c151-703">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c151-704">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-704">Requirements</span></span>

|<span data-ttu-id="8c151-705">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-705">Requirement</span></span>| <span data-ttu-id="8c151-706">値</span><span class="sxs-lookup"><span data-stu-id="8c151-706">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-707">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-707">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-708">1.3</span><span class="sxs-lookup"><span data-stu-id="8c151-708">1.3</span></span>|
|[<span data-ttu-id="8c151-709">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-709">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-710">制限あり</span><span class="sxs-lookup"><span data-stu-id="8c151-710">Restricted</span></span>|
|[<span data-ttu-id="8c151-711">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-711">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-712">作成</span><span class="sxs-lookup"><span data-stu-id="8c151-712">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="8c151-713">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="8c151-713">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="8c151-714">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-714">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8c151-715">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8c151-715">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8c151-716">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-716">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8c151-717">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="8c151-717">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="8c151-p138">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="8c151-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c151-721">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8c151-721">Parameters:</span></span>

| <span data-ttu-id="8c151-722">名前</span><span class="sxs-lookup"><span data-stu-id="8c151-722">Name</span></span> | <span data-ttu-id="8c151-723">型</span><span class="sxs-lookup"><span data-stu-id="8c151-723">Type</span></span> | <span data-ttu-id="8c151-724">属性</span><span class="sxs-lookup"><span data-stu-id="8c151-724">Attributes</span></span> | <span data-ttu-id="8c151-725">説明</span><span class="sxs-lookup"><span data-stu-id="8c151-725">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="8c151-726">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="8c151-726">String &#124; Object</span></span>| |<span data-ttu-id="8c151-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="8c151-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8c151-729">**または**</span><span class="sxs-lookup"><span data-stu-id="8c151-729">**OR**</span></span><br/><span data-ttu-id="8c151-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8c151-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="8c151-732">String</span><span class="sxs-lookup"><span data-stu-id="8c151-732">String</span></span> | <span data-ttu-id="8c151-733">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-733">&lt;optional&gt;</span></span> | <span data-ttu-id="8c151-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="8c151-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="8c151-736">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-736">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="8c151-737">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-737">&lt;optional&gt;</span></span> | <span data-ttu-id="8c151-738">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="8c151-738">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="8c151-739">String</span><span class="sxs-lookup"><span data-stu-id="8c151-739">String</span></span> | | <span data-ttu-id="8c151-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="8c151-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="8c151-742">String</span><span class="sxs-lookup"><span data-stu-id="8c151-742">String</span></span> | | <span data-ttu-id="8c151-743">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="8c151-743">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="8c151-744">String</span><span class="sxs-lookup"><span data-stu-id="8c151-744">String</span></span> | | <span data-ttu-id="8c151-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="8c151-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="8c151-747">ブール値</span><span class="sxs-lookup"><span data-stu-id="8c151-747">Boolean</span></span> | | <span data-ttu-id="8c151-p144">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="8c151-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="8c151-750">String</span><span class="sxs-lookup"><span data-stu-id="8c151-750">String</span></span> | | <span data-ttu-id="8c151-p145">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="8c151-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="8c151-754">function</span><span class="sxs-lookup"><span data-stu-id="8c151-754">function</span></span> | <span data-ttu-id="8c151-755">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-755">&lt;optional&gt;</span></span> | <span data-ttu-id="8c151-756">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8c151-757">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-757">Requirements</span></span>

|<span data-ttu-id="8c151-758">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-758">Requirement</span></span>| <span data-ttu-id="8c151-759">値</span><span class="sxs-lookup"><span data-stu-id="8c151-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-760">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-760">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-761">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-761">1.0</span></span>|
|[<span data-ttu-id="8c151-762">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-762">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-763">ReadItem</span></span>|
|[<span data-ttu-id="8c151-764">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-764">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-765">読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8c151-766">例</span><span class="sxs-lookup"><span data-stu-id="8c151-766">Examples</span></span>

<span data-ttu-id="8c151-767">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="8c151-767">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="8c151-768">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="8c151-768">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="8c151-769">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="8c151-769">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8c151-770">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="8c151-770">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="8c151-771">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="8c151-771">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="8c151-772">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="8c151-772">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="8c151-773">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="8c151-773">displayReplyForm(formData)</span></span>

<span data-ttu-id="8c151-774">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-774">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8c151-775">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8c151-775">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8c151-776">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-776">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8c151-777">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="8c151-777">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="8c151-p146">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="8c151-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c151-781">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8c151-781">Parameters:</span></span>

| <span data-ttu-id="8c151-782">名前</span><span class="sxs-lookup"><span data-stu-id="8c151-782">Name</span></span> | <span data-ttu-id="8c151-783">型</span><span class="sxs-lookup"><span data-stu-id="8c151-783">Type</span></span> | <span data-ttu-id="8c151-784">属性</span><span class="sxs-lookup"><span data-stu-id="8c151-784">Attributes</span></span> | <span data-ttu-id="8c151-785">説明</span><span class="sxs-lookup"><span data-stu-id="8c151-785">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="8c151-786">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="8c151-786">String &#124; Object</span></span>| | <span data-ttu-id="8c151-p147">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="8c151-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8c151-789">**または**</span><span class="sxs-lookup"><span data-stu-id="8c151-789">**OR**</span></span><br/><span data-ttu-id="8c151-p148">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8c151-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="8c151-792">String</span><span class="sxs-lookup"><span data-stu-id="8c151-792">String</span></span> | <span data-ttu-id="8c151-793">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-793">&lt;optional&gt;</span></span> | <span data-ttu-id="8c151-p149">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="8c151-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="8c151-796">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-796">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="8c151-797">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-797">&lt;optional&gt;</span></span> | <span data-ttu-id="8c151-798">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="8c151-798">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="8c151-799">String</span><span class="sxs-lookup"><span data-stu-id="8c151-799">String</span></span> | | <span data-ttu-id="8c151-p150">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="8c151-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="8c151-802">String</span><span class="sxs-lookup"><span data-stu-id="8c151-802">String</span></span> | | <span data-ttu-id="8c151-803">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="8c151-803">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="8c151-804">String</span><span class="sxs-lookup"><span data-stu-id="8c151-804">String</span></span> | | <span data-ttu-id="8c151-p151">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="8c151-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="8c151-807">ブール値</span><span class="sxs-lookup"><span data-stu-id="8c151-807">Boolean</span></span> | | <span data-ttu-id="8c151-p152">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="8c151-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="8c151-810">String</span><span class="sxs-lookup"><span data-stu-id="8c151-810">String</span></span> | | <span data-ttu-id="8c151-p153">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="8c151-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="8c151-814">function</span><span class="sxs-lookup"><span data-stu-id="8c151-814">function</span></span> | <span data-ttu-id="8c151-815">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-815">&lt;optional&gt;</span></span> | <span data-ttu-id="8c151-816">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-816">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8c151-817">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-817">Requirements</span></span>

|<span data-ttu-id="8c151-818">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-818">Requirement</span></span>| <span data-ttu-id="8c151-819">値</span><span class="sxs-lookup"><span data-stu-id="8c151-819">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-820">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-820">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-821">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-821">1.0</span></span>|
|[<span data-ttu-id="8c151-822">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-822">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-823">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-823">ReadItem</span></span>|
|[<span data-ttu-id="8c151-824">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-824">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-825">読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-825">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8c151-826">例</span><span class="sxs-lookup"><span data-stu-id="8c151-826">Examples</span></span>

<span data-ttu-id="8c151-827">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="8c151-827">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="8c151-828">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="8c151-828">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="8c151-829">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="8c151-829">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8c151-830">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="8c151-830">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="8c151-831">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="8c151-831">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="8c151-832">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="8c151-832">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="8c151-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="8c151-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="8c151-834">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="8c151-834">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="8c151-835">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8c151-835">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c151-836">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-836">Requirements</span></span>

|<span data-ttu-id="8c151-837">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-837">Requirement</span></span>| <span data-ttu-id="8c151-838">値</span><span class="sxs-lookup"><span data-stu-id="8c151-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-839">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-840">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-840">1.0</span></span>|
|[<span data-ttu-id="8c151-841">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-841">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-842">ReadItem</span></span>|
|[<span data-ttu-id="8c151-843">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-843">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-844">読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8c151-845">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8c151-845">Returns:</span></span>

<span data-ttu-id="8c151-846">型:[Entities](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="8c151-846">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="8c151-847">例</span><span class="sxs-lookup"><span data-stu-id="8c151-847">Example</span></span>

<span data-ttu-id="8c151-848">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="8c151-848">The following example accesses the contacts entities on the current item.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="8c151-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="8c151-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="8c151-850">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="8c151-850">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="8c151-851">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8c151-851">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c151-852">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8c151-852">Parameters:</span></span>

|<span data-ttu-id="8c151-853">名前</span><span class="sxs-lookup"><span data-stu-id="8c151-853">Name</span></span>| <span data-ttu-id="8c151-854">型</span><span class="sxs-lookup"><span data-stu-id="8c151-854">Type</span></span>| <span data-ttu-id="8c151-855">説明</span><span class="sxs-lookup"><span data-stu-id="8c151-855">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="8c151-856">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="8c151-856">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="8c151-857">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="8c151-857">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8c151-858">Requirements</span><span class="sxs-lookup"><span data-stu-id="8c151-858">Requirements</span></span>

|<span data-ttu-id="8c151-859">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-859">Requirement</span></span>| <span data-ttu-id="8c151-860">値</span><span class="sxs-lookup"><span data-stu-id="8c151-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-861">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-862">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-862">1.0</span></span>|
|[<span data-ttu-id="8c151-863">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-863">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-864">制限あり</span><span class="sxs-lookup"><span data-stu-id="8c151-864">Restricted</span></span>|
|[<span data-ttu-id="8c151-865">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-865">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-866">読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-866">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8c151-867">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8c151-867">Returns:</span></span>

<span data-ttu-id="8c151-868">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-868">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="8c151-869">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-869">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="8c151-870">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="8c151-870">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="8c151-871">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="8c151-871">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="8c151-872">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="8c151-872">Value of `entityType`</span></span> | <span data-ttu-id="8c151-873">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="8c151-873">Type of objects in returned array</span></span> | <span data-ttu-id="8c151-874">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="8c151-874">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="8c151-875">文字列</span><span class="sxs-lookup"><span data-stu-id="8c151-875">String</span></span> | <span data-ttu-id="8c151-876">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="8c151-876">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="8c151-877">連絡先</span><span class="sxs-lookup"><span data-stu-id="8c151-877">Contact</span></span> | <span data-ttu-id="8c151-878">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8c151-878">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="8c151-879">文字列</span><span class="sxs-lookup"><span data-stu-id="8c151-879">String</span></span> | <span data-ttu-id="8c151-880">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8c151-880">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="8c151-881">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="8c151-881">MeetingSuggestion</span></span> | <span data-ttu-id="8c151-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8c151-882">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="8c151-883">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="8c151-883">PhoneNumber</span></span> | <span data-ttu-id="8c151-884">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="8c151-884">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="8c151-885">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="8c151-885">TaskSuggestion</span></span> | <span data-ttu-id="8c151-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8c151-886">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="8c151-887">文字列</span><span class="sxs-lookup"><span data-stu-id="8c151-887">String</span></span> | <span data-ttu-id="8c151-888">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="8c151-888">**Restricted**</span></span> |

<span data-ttu-id="8c151-889">型:Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="8c151-889">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="8c151-890">例</span><span class="sxs-lookup"><span data-stu-id="8c151-890">Example</span></span>

<span data-ttu-id="8c151-891">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="8c151-891">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="8c151-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="8c151-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="8c151-893">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-893">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8c151-894">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8c151-894">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8c151-895">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-895">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c151-896">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8c151-896">Parameters:</span></span>

|<span data-ttu-id="8c151-897">名前</span><span class="sxs-lookup"><span data-stu-id="8c151-897">Name</span></span>| <span data-ttu-id="8c151-898">型</span><span class="sxs-lookup"><span data-stu-id="8c151-898">Type</span></span>| <span data-ttu-id="8c151-899">説明</span><span class="sxs-lookup"><span data-stu-id="8c151-899">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="8c151-900">String</span><span class="sxs-lookup"><span data-stu-id="8c151-900">String</span></span>|<span data-ttu-id="8c151-901">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="8c151-901">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8c151-902">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-902">Requirements</span></span>

|<span data-ttu-id="8c151-903">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-903">Requirement</span></span>| <span data-ttu-id="8c151-904">値</span><span class="sxs-lookup"><span data-stu-id="8c151-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-905">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-905">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-906">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-906">1.0</span></span>|
|[<span data-ttu-id="8c151-907">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-908">ReadItem</span></span>|
|[<span data-ttu-id="8c151-909">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-910">読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-910">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8c151-911">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8c151-911">Returns:</span></span>

<span data-ttu-id="8c151-p155">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="8c151-914">型:Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="8c151-914">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="8c151-915">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="8c151-915">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="8c151-916">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-916">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8c151-917">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8c151-917">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8c151-p156">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="8c151-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="8c151-921">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="8c151-921">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="8c151-922">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="8c151-922">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="8c151-p157">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="8c151-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c151-926">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-926">Requirements</span></span>

|<span data-ttu-id="8c151-927">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-927">Requirement</span></span>| <span data-ttu-id="8c151-928">値</span><span class="sxs-lookup"><span data-stu-id="8c151-928">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-929">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-929">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-930">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-930">1.0</span></span>|
|[<span data-ttu-id="8c151-931">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-931">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-932">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-932">ReadItem</span></span>|
|[<span data-ttu-id="8c151-933">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-933">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-934">読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-934">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8c151-935">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8c151-935">Returns:</span></span>

<span data-ttu-id="8c151-p158">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="8c151-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="8c151-938">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="8c151-938">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8c151-939">Object</span><span class="sxs-lookup"><span data-stu-id="8c151-939">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8c151-940">例</span><span class="sxs-lookup"><span data-stu-id="8c151-940">Example</span></span>

<span data-ttu-id="8c151-941">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="8c151-941">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="8c151-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="8c151-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="8c151-943">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-943">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8c151-944">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8c151-944">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8c151-945">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="8c151-945">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="8c151-p159">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="8c151-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c151-948">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8c151-948">Parameters:</span></span>

|<span data-ttu-id="8c151-949">名前</span><span class="sxs-lookup"><span data-stu-id="8c151-949">Name</span></span>| <span data-ttu-id="8c151-950">型</span><span class="sxs-lookup"><span data-stu-id="8c151-950">Type</span></span>| <span data-ttu-id="8c151-951">説明</span><span class="sxs-lookup"><span data-stu-id="8c151-951">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="8c151-952">String</span><span class="sxs-lookup"><span data-stu-id="8c151-952">String</span></span>|<span data-ttu-id="8c151-953">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="8c151-953">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8c151-954">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-954">Requirements</span></span>

|<span data-ttu-id="8c151-955">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-955">Requirement</span></span>| <span data-ttu-id="8c151-956">値</span><span class="sxs-lookup"><span data-stu-id="8c151-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-957">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-957">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-958">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-958">1.0</span></span>|
|[<span data-ttu-id="8c151-959">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-960">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-960">ReadItem</span></span>|
|[<span data-ttu-id="8c151-961">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-962">読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8c151-963">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8c151-963">Returns:</span></span>

<span data-ttu-id="8c151-964">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="8c151-964">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="8c151-965">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="8c151-965">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8c151-966">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="8c151-966">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8c151-967">例</span><span class="sxs-lookup"><span data-stu-id="8c151-967">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="8c151-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="8c151-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="8c151-969">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-969">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="8c151-p160">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c151-972">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8c151-972">Parameters:</span></span>

|<span data-ttu-id="8c151-973">名前</span><span class="sxs-lookup"><span data-stu-id="8c151-973">Name</span></span>| <span data-ttu-id="8c151-974">型</span><span class="sxs-lookup"><span data-stu-id="8c151-974">Type</span></span>| <span data-ttu-id="8c151-975">属性</span><span class="sxs-lookup"><span data-stu-id="8c151-975">Attributes</span></span>| <span data-ttu-id="8c151-976">説明</span><span class="sxs-lookup"><span data-stu-id="8c151-976">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="8c151-977">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8c151-977">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="8c151-p161">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="8c151-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="8c151-981">Object</span><span class="sxs-lookup"><span data-stu-id="8c151-981">Object</span></span>| <span data-ttu-id="8c151-982">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-982">&lt;optional&gt;</span></span>|<span data-ttu-id="8c151-983">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8c151-983">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8c151-984">Object</span><span class="sxs-lookup"><span data-stu-id="8c151-984">Object</span></span>| <span data-ttu-id="8c151-985">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-985">&lt;optional&gt;</span></span>|<span data-ttu-id="8c151-986">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8c151-986">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8c151-987">function</span><span class="sxs-lookup"><span data-stu-id="8c151-987">function</span></span>||<span data-ttu-id="8c151-988">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-988">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8c151-989">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="8c151-989">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="8c151-990">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="8c151-990">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8c151-991">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-991">Requirements</span></span>

|<span data-ttu-id="8c151-992">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-992">Requirement</span></span>| <span data-ttu-id="8c151-993">値</span><span class="sxs-lookup"><span data-stu-id="8c151-993">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-994">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-994">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-995">1.2</span><span class="sxs-lookup"><span data-stu-id="8c151-995">1.2</span></span>|
|[<span data-ttu-id="8c151-996">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-996">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-997">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8c151-997">ReadWriteItem</span></span>|
|[<span data-ttu-id="8c151-998">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-998">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-999">作成</span><span class="sxs-lookup"><span data-stu-id="8c151-999">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="8c151-1000">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8c151-1000">Returns:</span></span>

<span data-ttu-id="8c151-1001">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="8c151-1001">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="8c151-1002">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="8c151-1002">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8c151-1003">String</span><span class="sxs-lookup"><span data-stu-id="8c151-1003">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8c151-1004">例</span><span class="sxs-lookup"><span data-stu-id="8c151-1004">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="8c151-1005">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8c151-1005">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="8c151-1006">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="8c151-1006">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="8c151-p163">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="8c151-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c151-1010">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8c151-1010">Parameters:</span></span>

|<span data-ttu-id="8c151-1011">名前</span><span class="sxs-lookup"><span data-stu-id="8c151-1011">Name</span></span>| <span data-ttu-id="8c151-1012">型</span><span class="sxs-lookup"><span data-stu-id="8c151-1012">Type</span></span>| <span data-ttu-id="8c151-1013">属性</span><span class="sxs-lookup"><span data-stu-id="8c151-1013">Attributes</span></span>| <span data-ttu-id="8c151-1014">説明</span><span class="sxs-lookup"><span data-stu-id="8c151-1014">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="8c151-1015">function</span><span class="sxs-lookup"><span data-stu-id="8c151-1015">function</span></span>||<span data-ttu-id="8c151-1016">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8c151-1017">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-1017">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="8c151-1018">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="8c151-1018">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="8c151-1019">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8c151-1019">Object</span></span>| <span data-ttu-id="8c151-1020">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-1020">&lt;optional&gt;</span></span>|<span data-ttu-id="8c151-1021">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="8c151-1021">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="8c151-1022">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="8c151-1022">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8c151-1023">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-1023">Requirements</span></span>

|<span data-ttu-id="8c151-1024">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-1024">Requirement</span></span>| <span data-ttu-id="8c151-1025">値</span><span class="sxs-lookup"><span data-stu-id="8c151-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-1026">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-1026">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="8c151-1027">1.0</span></span>|
|[<span data-ttu-id="8c151-1028">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-1028">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c151-1029">ReadItem</span></span>|
|[<span data-ttu-id="8c151-1030">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-1030">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-1031">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8c151-1031">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-1032">例</span><span class="sxs-lookup"><span data-stu-id="8c151-1032">Example</span></span>

<span data-ttu-id="8c151-p166">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="8c151-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="8c151-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8c151-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="8c151-1037">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="8c151-1037">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="8c151-p167">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="8c151-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c151-1042">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8c151-1042">Parameters:</span></span>

|<span data-ttu-id="8c151-1043">名前</span><span class="sxs-lookup"><span data-stu-id="8c151-1043">Name</span></span>| <span data-ttu-id="8c151-1044">型</span><span class="sxs-lookup"><span data-stu-id="8c151-1044">Type</span></span>| <span data-ttu-id="8c151-1045">属性</span><span class="sxs-lookup"><span data-stu-id="8c151-1045">Attributes</span></span>| <span data-ttu-id="8c151-1046">説明</span><span class="sxs-lookup"><span data-stu-id="8c151-1046">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="8c151-1047">String</span><span class="sxs-lookup"><span data-stu-id="8c151-1047">String</span></span>||<span data-ttu-id="8c151-p168">削除する添付ファイルの識別子。文字列の最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="8c151-p168">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="8c151-1050">Object</span><span class="sxs-lookup"><span data-stu-id="8c151-1050">Object</span></span>| <span data-ttu-id="8c151-1051">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-1051">&lt;optional&gt;</span></span>|<span data-ttu-id="8c151-1052">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8c151-1052">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8c151-1053">Object</span><span class="sxs-lookup"><span data-stu-id="8c151-1053">Object</span></span>| <span data-ttu-id="8c151-1054">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-1054">&lt;optional&gt;</span></span>|<span data-ttu-id="8c151-1055">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8c151-1055">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8c151-1056">function</span><span class="sxs-lookup"><span data-stu-id="8c151-1056">function</span></span>| <span data-ttu-id="8c151-1057">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="8c151-1058">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8c151-1059">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="8c151-1059">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8c151-1060">エラー</span><span class="sxs-lookup"><span data-stu-id="8c151-1060">Errors</span></span>

| <span data-ttu-id="8c151-1061">エラー コード</span><span class="sxs-lookup"><span data-stu-id="8c151-1061">Error code</span></span> | <span data-ttu-id="8c151-1062">説明</span><span class="sxs-lookup"><span data-stu-id="8c151-1062">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="8c151-1063">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="8c151-1063">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8c151-1064">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-1064">Requirements</span></span>

|<span data-ttu-id="8c151-1065">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-1065">Requirement</span></span>| <span data-ttu-id="8c151-1066">値</span><span class="sxs-lookup"><span data-stu-id="8c151-1066">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-1067">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-1067">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-1068">1.1</span><span class="sxs-lookup"><span data-stu-id="8c151-1068">1.1</span></span>|
|[<span data-ttu-id="8c151-1069">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-1069">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-1070">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8c151-1070">ReadWriteItem</span></span>|
|[<span data-ttu-id="8c151-1071">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-1071">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-1072">作成</span><span class="sxs-lookup"><span data-stu-id="8c151-1072">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-1073">例</span><span class="sxs-lookup"><span data-stu-id="8c151-1073">Example</span></span>

<span data-ttu-id="8c151-1074">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="8c151-1074">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="8c151-1075">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="8c151-1075">saveAsync([options], callback)</span></span>

<span data-ttu-id="8c151-1076">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="8c151-1076">Asynchronously saves an item.</span></span>

<span data-ttu-id="8c151-p169">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-p169">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="8c151-1080">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="8c151-1080">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="8c151-1081">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-1081">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="8c151-p171">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-p171">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="8c151-1085">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="8c151-1085">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="8c151-1086">Mac Outlook では、新規作成モードの会議で `saveAsync` をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="8c151-1086">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="8c151-1087">Mac Outlook では、会議で `saveAsync` を呼び出すとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-1087">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="8c151-1088">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-1088">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c151-1089">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8c151-1089">Parameters:</span></span>

|<span data-ttu-id="8c151-1090">名前</span><span class="sxs-lookup"><span data-stu-id="8c151-1090">Name</span></span>| <span data-ttu-id="8c151-1091">型</span><span class="sxs-lookup"><span data-stu-id="8c151-1091">Type</span></span>| <span data-ttu-id="8c151-1092">属性</span><span class="sxs-lookup"><span data-stu-id="8c151-1092">Attributes</span></span>| <span data-ttu-id="8c151-1093">説明</span><span class="sxs-lookup"><span data-stu-id="8c151-1093">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="8c151-1094">Object</span><span class="sxs-lookup"><span data-stu-id="8c151-1094">Object</span></span>| <span data-ttu-id="8c151-1095">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="8c151-1096">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8c151-1096">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8c151-1097">Object</span><span class="sxs-lookup"><span data-stu-id="8c151-1097">Object</span></span>| <span data-ttu-id="8c151-1098">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="8c151-1099">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8c151-1099">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8c151-1100">function</span><span class="sxs-lookup"><span data-stu-id="8c151-1100">function</span></span>||<span data-ttu-id="8c151-1101">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8c151-1102">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-1102">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8c151-1103">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-1103">Requirements</span></span>

|<span data-ttu-id="8c151-1104">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-1104">Requirement</span></span>| <span data-ttu-id="8c151-1105">値</span><span class="sxs-lookup"><span data-stu-id="8c151-1105">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-1106">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-1106">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-1107">1.3</span><span class="sxs-lookup"><span data-stu-id="8c151-1107">1.3</span></span>|
|[<span data-ttu-id="8c151-1108">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-1108">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-1109">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8c151-1109">ReadWriteItem</span></span>|
|[<span data-ttu-id="8c151-1110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-1110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-1111">作成</span><span class="sxs-lookup"><span data-stu-id="8c151-1111">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8c151-1112">例</span><span class="sxs-lookup"><span data-stu-id="8c151-1112">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="8c151-p173">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="8c151-p173">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="8c151-1115">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="8c151-1115">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="8c151-1116">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="8c151-1116">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="8c151-p174">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="8c151-p174">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c151-1120">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8c151-1120">Parameters:</span></span>

|<span data-ttu-id="8c151-1121">名前</span><span class="sxs-lookup"><span data-stu-id="8c151-1121">Name</span></span>| <span data-ttu-id="8c151-1122">型</span><span class="sxs-lookup"><span data-stu-id="8c151-1122">Type</span></span>| <span data-ttu-id="8c151-1123">属性</span><span class="sxs-lookup"><span data-stu-id="8c151-1123">Attributes</span></span>| <span data-ttu-id="8c151-1124">説明</span><span class="sxs-lookup"><span data-stu-id="8c151-1124">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="8c151-1125">String</span><span class="sxs-lookup"><span data-stu-id="8c151-1125">String</span></span>||<span data-ttu-id="8c151-p175">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="8c151-p175">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="8c151-1129">Object</span><span class="sxs-lookup"><span data-stu-id="8c151-1129">Object</span></span>| <span data-ttu-id="8c151-1130">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="8c151-1131">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8c151-1131">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8c151-1132">Object</span><span class="sxs-lookup"><span data-stu-id="8c151-1132">Object</span></span>| <span data-ttu-id="8c151-1133">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="8c151-1134">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8c151-1134">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="8c151-1135">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8c151-1135">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="8c151-1136">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c151-1136">&lt;optional&gt;</span></span>|<span data-ttu-id="8c151-p176">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-p176">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="8c151-p177">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-p177">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="8c151-1141">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-1141">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="8c151-1142">function</span><span class="sxs-lookup"><span data-stu-id="8c151-1142">function</span></span>||<span data-ttu-id="8c151-1143">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8c151-1143">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8c151-1144">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-1144">Requirements</span></span>

|<span data-ttu-id="8c151-1145">要件</span><span class="sxs-lookup"><span data-stu-id="8c151-1145">Requirement</span></span>| <span data-ttu-id="8c151-1146">値</span><span class="sxs-lookup"><span data-stu-id="8c151-1146">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c151-1147">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8c151-1147">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c151-1148">1.2</span><span class="sxs-lookup"><span data-stu-id="8c151-1148">1.2</span></span>|
|[<span data-ttu-id="8c151-1149">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8c151-1149">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c151-1150">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8c151-1150">ReadWriteItem</span></span>|
|[<span data-ttu-id="8c151-1151">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8c151-1151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c151-1152">作成</span><span class="sxs-lookup"><span data-stu-id="8c151-1152">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8c151-1153">例</span><span class="sxs-lookup"><span data-stu-id="8c151-1153">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```