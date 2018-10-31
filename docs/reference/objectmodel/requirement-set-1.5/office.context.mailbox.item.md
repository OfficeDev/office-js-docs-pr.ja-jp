
# <a name="item"></a><span data-ttu-id="cb1ec-101">項目</span><span class="sxs-lookup"><span data-stu-id="cb1ec-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="cb1ec-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="cb1ec-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="cb1ec-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1ec-105">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-105">Requirements</span></span>

|<span data-ttu-id="cb1ec-106">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-106">Requirement</span></span>| <span data-ttu-id="cb1ec-107">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-109">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-109">1.0</span></span>|
|[<span data-ttu-id="cb1ec-110">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="cb1ec-111">Restricted</span></span>|
|[<span data-ttu-id="cb1ec-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-113">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="cb1ec-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="cb1ec-114">Members and methods</span></span>

| <span data-ttu-id="cb1ec-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-115">Member</span></span> | <span data-ttu-id="cb1ec-116">型</span><span class="sxs-lookup"><span data-stu-id="cb1ec-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="cb1ec-117">attachments</span><span class="sxs-lookup"><span data-stu-id="cb1ec-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails) | <span data-ttu-id="cb1ec-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-118">Member</span></span> |
| [<span data-ttu-id="cb1ec-119">bcc</span><span class="sxs-lookup"><span data-stu-id="cb1ec-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="cb1ec-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-120">Member</span></span> |
| [<span data-ttu-id="cb1ec-121">body</span><span class="sxs-lookup"><span data-stu-id="cb1ec-121">body</span></span>](#body-bodyjavascriptapioutlook15officebody) | <span data-ttu-id="cb1ec-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-122">Member</span></span> |
| [<span data-ttu-id="cb1ec-123">cc</span><span class="sxs-lookup"><span data-stu-id="cb1ec-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="cb1ec-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-124">Member</span></span> |
| [<span data-ttu-id="cb1ec-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="cb1ec-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="cb1ec-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-126">Member</span></span> |
| [<span data-ttu-id="cb1ec-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="cb1ec-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="cb1ec-128">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-128">Member</span></span> |
| [<span data-ttu-id="cb1ec-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="cb1ec-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="cb1ec-130">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-130">Member</span></span> |
| [<span data-ttu-id="cb1ec-131">end</span><span class="sxs-lookup"><span data-stu-id="cb1ec-131">end</span></span>](#end-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="cb1ec-132">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-132">Member</span></span> |
| [<span data-ttu-id="cb1ec-133">from</span><span class="sxs-lookup"><span data-stu-id="cb1ec-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="cb1ec-134">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-134">Member</span></span> |
| [<span data-ttu-id="cb1ec-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="cb1ec-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="cb1ec-136">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-136">Member</span></span> |
| [<span data-ttu-id="cb1ec-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="cb1ec-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="cb1ec-138">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-138">Member</span></span> |
| [<span data-ttu-id="cb1ec-139">itemId</span><span class="sxs-lookup"><span data-stu-id="cb1ec-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="cb1ec-140">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-140">Member</span></span> |
| [<span data-ttu-id="cb1ec-141">itemType</span><span class="sxs-lookup"><span data-stu-id="cb1ec-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) | <span data-ttu-id="cb1ec-142">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-142">Member</span></span> |
| [<span data-ttu-id="cb1ec-143">location</span><span class="sxs-lookup"><span data-stu-id="cb1ec-143">location</span></span>](#location-stringlocationjavascriptapioutlook15officelocation) | <span data-ttu-id="cb1ec-144">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-144">Member</span></span> |
| [<span data-ttu-id="cb1ec-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="cb1ec-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="cb1ec-146">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-146">Member</span></span> |
| [<span data-ttu-id="cb1ec-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="cb1ec-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages) | <span data-ttu-id="cb1ec-148">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-148">Member</span></span> |
| [<span data-ttu-id="cb1ec-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="cb1ec-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="cb1ec-150">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-150">Member</span></span> |
| [<span data-ttu-id="cb1ec-151">主催者</span><span class="sxs-lookup"><span data-stu-id="cb1ec-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="cb1ec-152">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-152">Member</span></span> |
| [<span data-ttu-id="cb1ec-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="cb1ec-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="cb1ec-154">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-154">Member</span></span> |
| [<span data-ttu-id="cb1ec-155">送り主</span><span class="sxs-lookup"><span data-stu-id="cb1ec-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="cb1ec-156">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-156">Member</span></span> |
| [<span data-ttu-id="cb1ec-157">開始</span><span class="sxs-lookup"><span data-stu-id="cb1ec-157">start</span></span>](#start-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="cb1ec-158">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-158">Member</span></span> |
| [<span data-ttu-id="cb1ec-159">件名</span><span class="sxs-lookup"><span data-stu-id="cb1ec-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook15officesubject) | <span data-ttu-id="cb1ec-160">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-160">Member</span></span> |
| [<span data-ttu-id="cb1ec-161">宛先</span><span class="sxs-lookup"><span data-stu-id="cb1ec-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="cb1ec-162">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-162">Member</span></span> |
| [<span data-ttu-id="cb1ec-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="cb1ec-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="cb1ec-164">メソッド</span><span class="sxs-lookup"><span data-stu-id="cb1ec-164">Method</span></span> |
| [<span data-ttu-id="cb1ec-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="cb1ec-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="cb1ec-166">メソッド</span><span class="sxs-lookup"><span data-stu-id="cb1ec-166">Method</span></span> |
| [<span data-ttu-id="cb1ec-167">終了</span><span class="sxs-lookup"><span data-stu-id="cb1ec-167">close</span></span>](#close) | <span data-ttu-id="cb1ec-168">メソッド</span><span class="sxs-lookup"><span data-stu-id="cb1ec-168">Method</span></span> |
| [<span data-ttu-id="cb1ec-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="cb1ec-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="cb1ec-170">メソッド</span><span class="sxs-lookup"><span data-stu-id="cb1ec-170">Method</span></span> |
| [<span data-ttu-id="cb1ec-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="cb1ec-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="cb1ec-172">メソッド</span><span class="sxs-lookup"><span data-stu-id="cb1ec-172">Method</span></span> |
| [<span data-ttu-id="cb1ec-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="cb1ec-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook15officeentities) | <span data-ttu-id="cb1ec-174">メソッド</span><span class="sxs-lookup"><span data-stu-id="cb1ec-174">Method</span></span> |
| [<span data-ttu-id="cb1ec-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="cb1ec-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="cb1ec-176">メソッド</span><span class="sxs-lookup"><span data-stu-id="cb1ec-176">Method</span></span> |
| [<span data-ttu-id="cb1ec-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="cb1ec-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="cb1ec-178">メソッド</span><span class="sxs-lookup"><span data-stu-id="cb1ec-178">Method</span></span> |
| [<span data-ttu-id="cb1ec-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="cb1ec-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="cb1ec-180">メソッド</span><span class="sxs-lookup"><span data-stu-id="cb1ec-180">Method</span></span> |
| [<span data-ttu-id="cb1ec-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="cb1ec-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="cb1ec-182">メソッド</span><span class="sxs-lookup"><span data-stu-id="cb1ec-182">Method</span></span> |
| [<span data-ttu-id="cb1ec-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="cb1ec-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="cb1ec-184">メソッド</span><span class="sxs-lookup"><span data-stu-id="cb1ec-184">Method</span></span> |
| [<span data-ttu-id="cb1ec-185">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="cb1ec-185">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="cb1ec-186">メソッド</span><span class="sxs-lookup"><span data-stu-id="cb1ec-186">Method</span></span> |
| [<span data-ttu-id="cb1ec-187">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="cb1ec-187">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="cb1ec-188">メソッド</span><span class="sxs-lookup"><span data-stu-id="cb1ec-188">Method</span></span> |
| [<span data-ttu-id="cb1ec-189">saveAsync</span><span class="sxs-lookup"><span data-stu-id="cb1ec-189">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="cb1ec-190">メソッド</span><span class="sxs-lookup"><span data-stu-id="cb1ec-190">Method</span></span> |
| [<span data-ttu-id="cb1ec-191">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="cb1ec-191">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="cb1ec-192">メソッド</span><span class="sxs-lookup"><span data-stu-id="cb1ec-192">Method</span></span> |

### <a name="example"></a><span data-ttu-id="cb1ec-193">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-193">Example</span></span>

<span data-ttu-id="cb1ec-194">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-194">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="cb1ec-195">メンバー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-195">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="cb1ec-196">添付ファイル：配列.<[添付ファイルの詳細 ](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="cb1ec-196">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="cb1ec-p102">項目の添付ファイルの配列を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="cb1ec-199">潜在的なセキュリティ問題により特定の種類のファイルは、Outlookでブロックされ、したがって戻ってきません。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-199">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="cb1ec-200">詳細については、[「Outlook でブロックされた添付ファイル」](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)をご覧下さい。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-200">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-201">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-201">Type:</span></span>

*   <span data-ttu-id="cb1ec-202">配列.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="cb1ec-202">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1ec-203">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-203">Requirements</span></span>

|<span data-ttu-id="cb1ec-204">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-204">Requirement</span></span>| <span data-ttu-id="cb1ec-205">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-206">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-207">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-207">1.0</span></span>|
|[<span data-ttu-id="cb1ec-208">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-209">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-210">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-211">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-211">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-212">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-212">Example</span></span>

<span data-ttu-id="cb1ec-213">次のコードでは、現在の項目にあるすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-213">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="cb1ec-214">bcc:[受信者](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-214">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="cb1ec-215">メッセージの BCC (ブラインド カーボン コピー) 列上の 受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-215">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="cb1ec-216">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-216">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-217">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-217">Type:</span></span>

*   [<span data-ttu-id="cb1ec-218">受信者</span><span class="sxs-lookup"><span data-stu-id="cb1ec-218">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="cb1ec-219">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-219">Requirements</span></span>

|<span data-ttu-id="cb1ec-220">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-220">Requirement</span></span>| <span data-ttu-id="cb1ec-221">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-222">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-222">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-223">1.1</span><span class="sxs-lookup"><span data-stu-id="cb1ec-223">1.1</span></span>|
|[<span data-ttu-id="cb1ec-224">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-224">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-225">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-226">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-226">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-227">新規作成</span><span class="sxs-lookup"><span data-stu-id="cb1ec-227">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-228">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-228">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="cb1ec-229">本文:[本文](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-229">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="cb1ec-230">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-230">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-231">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-231">Type:</span></span>

*   [<span data-ttu-id="cb1ec-232">本文</span><span class="sxs-lookup"><span data-stu-id="cb1ec-232">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="cb1ec-233">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-233">Requirements</span></span>

|<span data-ttu-id="cb1ec-234">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-234">Requirement</span></span>| <span data-ttu-id="cb1ec-235">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-236">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-237">1.1</span><span class="sxs-lookup"><span data-stu-id="cb1ec-237">1.1</span></span>|
|[<span data-ttu-id="cb1ec-238">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-238">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-239">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-240">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-240">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-241">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-241">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="cb1ec-242">cc: 配列。 <[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-242">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="cb1ec-243">メッセージの CC (カーボン コピー) 受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-243">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="cb1ec-244">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-244">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cb1ec-245">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-245">Read mode</span></span>

<span data-ttu-id="cb1ec-p106">`cc`プロパティは、メッセージの**CC**列にある各受信者一覧の`EmailAddressDetails`オブジェクトを含む配列を返します。コレクションは最大100個のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="cb1ec-248">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-248">Compose mode</span></span>

<span data-ttu-id="cb1ec-249">`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-249">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-250">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-250">Type:</span></span>

*   <span data-ttu-id="cb1ec-251">配列 。<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-251">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1ec-252">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-252">Requirements</span></span>

|<span data-ttu-id="cb1ec-253">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-253">Requirement</span></span>| <span data-ttu-id="cb1ec-254">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-254">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-255">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-255">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-256">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-256">1.0</span></span>|
|[<span data-ttu-id="cb1ec-257">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-257">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-258">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-258">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-259">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-259">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-260">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-260">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-261">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-261">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="cb1ec-262">（空白が可能）conversationId：文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-262">(nullable) conversationId :String</span></span>

<span data-ttu-id="cb1ec-263">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-263">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="cb1ec-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="cb1ec-p108">作成フォームの新しいアイテムに対してこのプロパティの Null を取得します。ユーザーが件名を設定し項目を保存する場合、`conversationId`プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-268">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-268">Type:</span></span>

*   <span data-ttu-id="cb1ec-269">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-269">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1ec-270">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-270">Requirements</span></span>

|<span data-ttu-id="cb1ec-271">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-271">Requirement</span></span>| <span data-ttu-id="cb1ec-272">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-273">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-274">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-274">1.0</span></span>|
|[<span data-ttu-id="cb1ec-275">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-276">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-277">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-278">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-278">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="cb1ec-279">dateTimeCreated: 日付</span><span class="sxs-lookup"><span data-stu-id="cb1ec-279">dateTimeCreated :Date</span></span>

<span data-ttu-id="cb1ec-p109">アイテムが作成された日時を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-282">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-282">Type:</span></span>

*   <span data-ttu-id="cb1ec-283">日付</span><span class="sxs-lookup"><span data-stu-id="cb1ec-283">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1ec-284">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-284">Requirements</span></span>

|<span data-ttu-id="cb1ec-285">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-285">Requirement</span></span>| <span data-ttu-id="cb1ec-286">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-287">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-287">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-288">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-288">1.0</span></span>|
|[<span data-ttu-id="cb1ec-289">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-289">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-290">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-290">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-291">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-291">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-292">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-292">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-293">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-293">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="cb1ec-294">dateTimeModified: 日付</span><span class="sxs-lookup"><span data-stu-id="cb1ec-294">dateTimeModified :Date</span></span>

<span data-ttu-id="cb1ec-p110">アイテムが最後に変更された日時を取得します。読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="cb1ec-297">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-297">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-298">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-298">Type:</span></span>

*   <span data-ttu-id="cb1ec-299">日付</span><span class="sxs-lookup"><span data-stu-id="cb1ec-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1ec-300">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-300">Requirements</span></span>

|<span data-ttu-id="cb1ec-301">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-301">Requirement</span></span>| <span data-ttu-id="cb1ec-302">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-303">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-304">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-304">1.0</span></span>|
|[<span data-ttu-id="cb1ec-305">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-306">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-307">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-308">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-309">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-309">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="cb1ec-310">end:日付|[時間](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-310">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="cb1ec-311">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-311">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="cb1ec-p111">`end`プロパティは、協定世界時 (UTC) 形式の時刻値として表示されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cb1ec-314">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-314">Read mode</span></span>

<span data-ttu-id="cb1ec-315">`end`プロパティは`Date`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-315">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="cb1ec-316">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-316">Compose mode</span></span>

<span data-ttu-id="cb1ec-317">`end`プロパティは`Time`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-317">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="cb1ec-318">[ `Time.setAsync` ](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-)   メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-318">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-319">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-319">Type:</span></span>

*   <span data-ttu-id="cb1ec-320">日付| [時間](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-320">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1ec-321">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-321">Requirements</span></span>

|<span data-ttu-id="cb1ec-322">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-322">Requirement</span></span>| <span data-ttu-id="cb1ec-323">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-324">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-325">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-325">1.0</span></span>|
|[<span data-ttu-id="cb1ec-326">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-327">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-328">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-329">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-329">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-330">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-330">Example</span></span>

<span data-ttu-id="cb1ec-331">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-331">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="cb1ec-332">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-332">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="cb1ec-p112">メッセージの送信者の電子メール アドレスを取得します。読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="cb1ec-p113">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) プロパティは同一人物を表します。その場合、`from` プロパティは委任者を、送信者プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="cb1ec-337">`from` プロパティ内の`EmailAddressDetails` オブジェクトの`recipientType`  プロパティは、 `undefined`です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-337">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-338">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-338">Type:</span></span>

*   [<span data-ttu-id="cb1ec-339">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="cb1ec-339">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="cb1ec-340">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-340">Requirements</span></span>

|<span data-ttu-id="cb1ec-341">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-341">Requirement</span></span>| <span data-ttu-id="cb1ec-342">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-342">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-343">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-343">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-344">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-344">1.0</span></span>|
|[<span data-ttu-id="cb1ec-345">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-345">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-346">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-346">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-347">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-347">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-348">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-348">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="cb1ec-349">internetMessageId: 文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-349">internetMessageId :String</span></span>

<span data-ttu-id="cb1ec-p114">電子メール メッセージのインターネット メッセージ 識別子を取得します。読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-352">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-352">Type:</span></span>

*   <span data-ttu-id="cb1ec-353">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-353">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1ec-354">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-354">Requirements</span></span>

|<span data-ttu-id="cb1ec-355">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-355">Requirement</span></span>| <span data-ttu-id="cb1ec-356">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-357">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-358">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-358">1.0</span></span>|
|[<span data-ttu-id="cb1ec-359">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-359">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-360">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-361">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-361">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-362">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-362">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-363">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-363">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="cb1ec-364">itemClass: 文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-364">itemClass :String</span></span>

<span data-ttu-id="cb1ec-p115">選択された項目の Exchange Web サービスの項目 クラスを取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="cb1ec-p116">`itemClass` プロパティは、選択したアイテムのメッセージ クラスを指定します。次は、メッセージまたは予定アイテムの既定のメッセージ クラスを示しています。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="cb1ec-369">種類</span><span class="sxs-lookup"><span data-stu-id="cb1ec-369">Type</span></span> | <span data-ttu-id="cb1ec-370">説明</span><span class="sxs-lookup"><span data-stu-id="cb1ec-370">Description</span></span> | <span data-ttu-id="cb1ec-371">項目のクラス</span><span class="sxs-lookup"><span data-stu-id="cb1ec-371">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="cb1ec-372">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="cb1ec-372">Appointment items</span></span> | <span data-ttu-id="cb1ec-373">これらは、アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurence` の予定表アイテムです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-373">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="cb1ec-374">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="cb1ec-374">Message items</span></span> | <span data-ttu-id="cb1ec-375">これには、基本のメッセージ クラス として `IPM.Schedule.Meeting`  を使用する、既定のメッセージ クラス `IPM.Note`  会議出席依頼、返信および取り消しを持つ電子メール メッセージが含まれます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-375">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="cb1ec-376">既定のメッセージ クラスを拡張したカスタム メッセージ クラス 、たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など)を作成できます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-376">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-377">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-377">Type:</span></span>

*   <span data-ttu-id="cb1ec-378">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-378">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1ec-379">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-379">Requirements</span></span>

|<span data-ttu-id="cb1ec-380">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-380">Requirement</span></span>| <span data-ttu-id="cb1ec-381">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-381">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-382">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-382">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-383">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-383">1.0</span></span>|
|[<span data-ttu-id="cb1ec-384">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-384">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-385">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-386">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-386">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-387">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-387">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-388">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-388">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="cb1ec-389">（空白が可能） itemId ：文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-389">(nullable) itemId :String</span></span>

<span data-ttu-id="cb1ec-p117">現在の項目の Exchange Web サービスのアイテム識別子を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="cb1ec-392">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-392">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="cb1ec-393">`itemId` プロパティは、Outlook の Entry ID または Outlook の REST API によって使用される ID と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-393">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="cb1ec-394">この値を使用して REST API の呼び出しを行う前に [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-394">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="cb1ec-395">詳細については、 [Outlook アドインから Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-395">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="cb1ec-p119">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-398">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-398">Type:</span></span>

*   <span data-ttu-id="cb1ec-399">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-399">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1ec-400">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-400">Requirements</span></span>

|<span data-ttu-id="cb1ec-401">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-401">Requirement</span></span>| <span data-ttu-id="cb1ec-402">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-403">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-403">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-404">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-404">1.0</span></span>|
|[<span data-ttu-id="cb1ec-405">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-405">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-406">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-407">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-407">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-408">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-408">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-409">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-409">Example</span></span>

<span data-ttu-id="cb1ec-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="cb1ec-412">itemType:[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-412">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="cb1ec-413">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-413">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="cb1ec-414">`itemType`プロパティは、`ItemType`列挙値の 1 つを返します。これは`item`オブジェクト インスタンスがメッセージまたは予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-414">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-415">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-415">Type:</span></span>

*   [<span data-ttu-id="cb1ec-416">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="cb1ec-416">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="cb1ec-417">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-417">Requirements</span></span>

|<span data-ttu-id="cb1ec-418">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-418">Requirement</span></span>| <span data-ttu-id="cb1ec-419">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-420">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-420">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-421">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-421">1.0</span></span>|
|[<span data-ttu-id="cb1ec-422">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-422">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-423">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-424">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-424">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-425">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-425">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-426">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-426">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="cb1ec-427">位置: 文字列|[](/javascript/api/outlook_1_5/office.location)位置</span><span class="sxs-lookup"><span data-stu-id="cb1ec-427">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="cb1ec-428">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-428">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cb1ec-429">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-429">Read mode</span></span>

<span data-ttu-id="cb1ec-430">`location` プロパティは、予定の場所を含む文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-430">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="cb1ec-431">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-431">Compose mode</span></span>

<span data-ttu-id="cb1ec-432">`location` プロパティは、予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-432">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-433">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-433">Type:</span></span>

*   <span data-ttu-id="cb1ec-434">文字列 | [場所](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-434">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1ec-435">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-435">Requirements</span></span>

|<span data-ttu-id="cb1ec-436">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-436">Requirement</span></span>| <span data-ttu-id="cb1ec-437">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-437">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-438">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-438">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-439">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-439">1.0</span></span>|
|[<span data-ttu-id="cb1ec-440">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-440">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-441">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-442">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-442">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-443">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-443">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-444">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-444">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="cb1ec-445">normalizedSubject: 文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-445">normalizedSubject :String</span></span>

<span data-ttu-id="cb1ec-p121">すべてのプレフィックス (`RE:` や `FWD:` を含む) が削除された項目の件名を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="cb1ec-p122">normalizedSubject プロパティは、電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたアイテムの件名を取得します。プレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-450">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-450">Type:</span></span>

*   <span data-ttu-id="cb1ec-451">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-451">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1ec-452">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-452">Requirements</span></span>

|<span data-ttu-id="cb1ec-453">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-453">Requirement</span></span>| <span data-ttu-id="cb1ec-454">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-455">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-455">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-456">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-456">1.0</span></span>|
|[<span data-ttu-id="cb1ec-457">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-457">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-458">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-459">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-459">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-460">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-460">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-461">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-461">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="cb1ec-462">notificationMessages:[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-462">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="cb1ec-463">項目の通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-463">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-464">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-464">Type:</span></span>

*   [<span data-ttu-id="cb1ec-465">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="cb1ec-465">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="cb1ec-466">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-466">Requirements</span></span>

|<span data-ttu-id="cb1ec-467">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-467">Requirement</span></span>| <span data-ttu-id="cb1ec-468">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-469">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-470">1.3</span><span class="sxs-lookup"><span data-stu-id="cb1ec-470">1.3</span></span>|
|[<span data-ttu-id="cb1ec-471">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-471">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-472">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-473">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-473">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-474">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-474">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="cb1ec-475">optionalAttendees: 配列.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-475">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="cb1ec-476">イベントの任意の出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-476">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="cb1ec-477">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-477">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cb1ec-478">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-478">Read mode</span></span>

<span data-ttu-id="cb1ec-479">`optionalAttendees`プロパティは、会議への各任意出席者の`EmailAddressDetails`オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-479">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="cb1ec-480">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-480">Compose mode</span></span>

<span data-ttu-id="cb1ec-481">`optionalAttendees`プロパティは会議への任意出席者を取得または設定するためのメソッドを提供する`Recipients`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-481">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-482">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-482">Type:</span></span>

*   <span data-ttu-id="cb1ec-483">配列 。<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-483">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1ec-484">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-484">Requirements</span></span>

|<span data-ttu-id="cb1ec-485">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-485">Requirement</span></span>| <span data-ttu-id="cb1ec-486">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-487">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-487">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-488">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-488">1.0</span></span>|
|[<span data-ttu-id="cb1ec-489">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-489">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-490">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-491">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-491">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-492">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-492">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-493">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-493">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="cb1ec-494">開催者:[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-494">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="cb1ec-p124">指定の会議の開催者の電子メール アドレスを取得します。読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-497">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-497">Type:</span></span>

*   [<span data-ttu-id="cb1ec-498">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="cb1ec-498">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="cb1ec-499">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-499">Requirements</span></span>

|<span data-ttu-id="cb1ec-500">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-500">Requirement</span></span>| <span data-ttu-id="cb1ec-501">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-502">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-503">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-503">1.0</span></span>|
|[<span data-ttu-id="cb1ec-504">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-505">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-506">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-507">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-507">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-508">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-508">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="cb1ec-509">requiredAttendees: 配列 。<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_5/office.recipients)。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-509">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="cb1ec-510">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-510">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="cb1ec-511">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cb1ec-512">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-512">Read mode</span></span>

<span data-ttu-id="cb1ec-513">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-513">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="cb1ec-514">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-514">Compose mode</span></span>

<span data-ttu-id="cb1ec-515">`requiredAttendees` プロパティは会議への必須出席者を取得または設定するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-515">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-516">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-516">Type:</span></span>

*   <span data-ttu-id="cb1ec-517">配列 。<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1ec-518">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-518">Requirements</span></span>

|<span data-ttu-id="cb1ec-519">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-519">Requirement</span></span>| <span data-ttu-id="cb1ec-520">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-521">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-522">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-522">1.0</span></span>|
|[<span data-ttu-id="cb1ec-523">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-524">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-525">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-526">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-527">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-527">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="cb1ec-528">送信者:[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-528">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="cb1ec-p126">電子メール送信者のメールアドレスを取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="cb1ec-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。その場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="cb1ec-533">`sender`プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType`プロパティは、`undefined`です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-533">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-534">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-534">Type:</span></span>

*   [<span data-ttu-id="cb1ec-535">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="cb1ec-535">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="cb1ec-536">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-536">Requirements</span></span>

|<span data-ttu-id="cb1ec-537">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-537">Requirement</span></span>| <span data-ttu-id="cb1ec-538">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-539">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-540">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-540">1.0</span></span>|
|[<span data-ttu-id="cb1ec-541">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-541">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-542">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-543">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-543">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-544">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-544">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-545">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-545">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="cb1ec-546">開始: 日付 |[  時間](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-546">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="cb1ec-547">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-547">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="cb1ec-p128">`start`プロパティは、協定世界時 (UTC) 形式の時刻値として表示されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime)メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cb1ec-550">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-550">Read mode</span></span>

<span data-ttu-id="cb1ec-551">`start`プロパティは`Date`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-551">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="cb1ec-552">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-552">Compose mode</span></span>

<span data-ttu-id="cb1ec-553">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-553">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="cb1ec-554">[ `Time.setAsync` ](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-554">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-555">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-555">Type:</span></span>

*   <span data-ttu-id="cb1ec-556">日付| [時間](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-556">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1ec-557">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-557">Requirements</span></span>

|<span data-ttu-id="cb1ec-558">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-558">Requirement</span></span>| <span data-ttu-id="cb1ec-559">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-559">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-560">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-560">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-561">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-561">1.0</span></span>|
|[<span data-ttu-id="cb1ec-562">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-562">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-563">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-563">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-564">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-564">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-565">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-565">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-566">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-566">Example</span></span>

<span data-ttu-id="cb1ec-567">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-567">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="cb1ec-568">件名: 文字列 | [件名](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-568">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="cb1ec-569">アイテムの件名フィールドに表示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-569">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="cb1ec-570">`subject`プロパティは、電子メールサーバーから送信された項目の全件名を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-570">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cb1ec-571">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-571">Read mode</span></span>

<span data-ttu-id="cb1ec-p129">`subject`プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string)プロパティを使用して、行間にある`RE:`や`FW:`のなどのすべてのプレフィックスを削除した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="cb1ec-574">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-574">Compose mode</span></span>

<span data-ttu-id="cb1ec-575">`subject`プロパティは、件名を取得または設定するためのメソッドを提供する`Subject`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-575">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="cb1ec-576">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-576">Type:</span></span>

*   <span data-ttu-id="cb1ec-577">文字列 | [件名](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-577">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1ec-578">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-578">Requirements</span></span>

|<span data-ttu-id="cb1ec-579">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-579">Requirement</span></span>| <span data-ttu-id="cb1ec-580">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-581">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-581">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-582">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-582">1.0</span></span>|
|[<span data-ttu-id="cb1ec-583">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-583">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-584">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-585">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-585">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-586">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-586">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="cb1ec-587">to: 配列。 <[  EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[ 受信者](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-587">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="cb1ec-588">メッセージの **宛先**列にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-588">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="cb1ec-589">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-589">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cb1ec-590">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-590">Read mode</span></span>

<span data-ttu-id="cb1ec-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="cb1ec-593">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-593">Compose mode</span></span>

<span data-ttu-id="cb1ec-594">`to` プロパティは、メッセージの **To** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-594">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="cb1ec-595">型:</span><span class="sxs-lookup"><span data-stu-id="cb1ec-595">Type:</span></span>

*   <span data-ttu-id="cb1ec-596">配列 。<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-596">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1ec-597">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-597">Requirements</span></span>

|<span data-ttu-id="cb1ec-598">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-598">Requirement</span></span>| <span data-ttu-id="cb1ec-599">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-599">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-600">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-600">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-601">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-601">1.0</span></span>|
|[<span data-ttu-id="cb1ec-602">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-602">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-603">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-603">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-604">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-604">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-605">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-605">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-606">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-606">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="cb1ec-607">メソッド</span><span class="sxs-lookup"><span data-stu-id="cb1ec-607">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="cb1ec-608">addFileAttachmentAsync(uri, attachmentName, [オプション], [コールバック])</span><span class="sxs-lookup"><span data-stu-id="cb1ec-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="cb1ec-609">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="cb1ec-610">`addFileAttachmentAsync`メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内の項目に添付します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="cb1ec-611">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cb1ec-612">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="cb1ec-612">Parameters:</span></span>

|<span data-ttu-id="cb1ec-613">名前</span><span class="sxs-lookup"><span data-stu-id="cb1ec-613">Name</span></span>| <span data-ttu-id="cb1ec-614">種類</span><span class="sxs-lookup"><span data-stu-id="cb1ec-614">Type</span></span>| <span data-ttu-id="cb1ec-615">属性</span><span class="sxs-lookup"><span data-stu-id="cb1ec-615">Attributes</span></span>| <span data-ttu-id="cb1ec-616">説明</span><span class="sxs-lookup"><span data-stu-id="cb1ec-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="cb1ec-617">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-617">String</span></span>||<span data-ttu-id="cb1ec-p132">メッセージまたは予定に添付するファイルの場所を示す URIです。 2048 文字以内で入力してください。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="cb1ec-620">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-620">String</span></span>||<span data-ttu-id="cb1ec-p133">アップロード中に表示される添付ファイルがそのファイルの名前です。 255 文字以内で入力してください。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="cb1ec-623">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cb1ec-623">Object</span></span>| <span data-ttu-id="cb1ec-624">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-624">&lt;optional&gt;</span></span>|<span data-ttu-id="cb1ec-625">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-625">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="cb1ec-626">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cb1ec-626">Object</span></span> | <span data-ttu-id="cb1ec-627">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-627">&lt;optional&gt;</span></span> | <span data-ttu-id="cb1ec-628">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="cb1ec-629">ブール値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-629">Boolean</span></span> | <span data-ttu-id="cb1ec-630">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-630">&lt;optional&gt;</span></span> | <span data-ttu-id="cb1ec-631">`true`の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-631">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="cb1ec-632">関数</span><span class="sxs-lookup"><span data-stu-id="cb1ec-632">function</span></span>| <span data-ttu-id="cb1ec-633">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-633">&lt;optional&gt;</span></span>|<span data-ttu-id="cb1ec-634">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="cb1ec-635">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-635">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="cb1ec-636">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-636">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="cb1ec-637">エラー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-637">Errors</span></span>

| <span data-ttu-id="cb1ec-638">エラー コード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-638">Error code</span></span> | <span data-ttu-id="cb1ec-639">説明</span><span class="sxs-lookup"><span data-stu-id="cb1ec-639">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="cb1ec-640">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-640">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="cb1ec-641">許可されていない拡張子付きの添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-641">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="cb1ec-642">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-642">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="cb1ec-643">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-643">Requirements</span></span>

|<span data-ttu-id="cb1ec-644">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-644">Requirement</span></span>| <span data-ttu-id="cb1ec-645">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-645">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-646">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-646">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-647">1.1</span><span class="sxs-lookup"><span data-stu-id="cb1ec-647">1.1</span></span>|
|[<span data-ttu-id="cb1ec-648">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-648">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-649">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-649">ReadWriteItem</span></span>|
|[<span data-ttu-id="cb1ec-650">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-650">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-651">新規作成</span><span class="sxs-lookup"><span data-stu-id="cb1ec-651">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="cb1ec-652">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-652">Examples</span></span>

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

<span data-ttu-id="cb1ec-653">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-653">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="cb1ec-654">addItemAttachmentAsync (itemd、attachmentName、[オプション]、[コールバック])</span><span class="sxs-lookup"><span data-stu-id="cb1ec-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="cb1ec-655">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-655">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="cb1ec-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つ項目を新規作成フォーム内の項目に添付します。コールバック メソッドを指定する場合、`asyncResult` というパラメータがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、または項目を添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメータを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="cb1ec-659">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-659">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="cb1ec-660">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドは項目を、編集中以外の項目に添付できますが、これはサポートされておらず、推奨されていません。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-660">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cb1ec-661">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="cb1ec-661">Parameters:</span></span>

|<span data-ttu-id="cb1ec-662">名前</span><span class="sxs-lookup"><span data-stu-id="cb1ec-662">Name</span></span>| <span data-ttu-id="cb1ec-663">種類</span><span class="sxs-lookup"><span data-stu-id="cb1ec-663">Type</span></span>| <span data-ttu-id="cb1ec-664">属性</span><span class="sxs-lookup"><span data-stu-id="cb1ec-664">Attributes</span></span>| <span data-ttu-id="cb1ec-665">説明</span><span class="sxs-lookup"><span data-stu-id="cb1ec-665">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="cb1ec-666">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-666">String</span></span>||<span data-ttu-id="cb1ec-p135">添付するアイテムの Exchange 識別子です。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="cb1ec-669">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-669">String</span></span>||<span data-ttu-id="cb1ec-p136">添付するアイテムの件名です。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="cb1ec-672">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cb1ec-672">Object</span></span>| <span data-ttu-id="cb1ec-673">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-673">&lt;optional&gt;</span></span>|<span data-ttu-id="cb1ec-674">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-674">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="cb1ec-675">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cb1ec-675">Object</span></span>| <span data-ttu-id="cb1ec-676">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-676">&lt;optional&gt;</span></span>|<span data-ttu-id="cb1ec-677">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-677">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="cb1ec-678">関数</span><span class="sxs-lookup"><span data-stu-id="cb1ec-678">function</span></span>| <span data-ttu-id="cb1ec-679">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-679">&lt;optional&gt;</span></span>|<span data-ttu-id="cb1ec-680">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-680">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="cb1ec-681">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-681">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="cb1ec-682">添付ファイルの追加に失敗した場合、 エラーの説明を提供する`asyncResult` オブジェクトが `Error` オブジェクトに含まれます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-682">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="cb1ec-683">エラー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-683">Errors</span></span>

| <span data-ttu-id="cb1ec-684">エラー コード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-684">Error code</span></span> | <span data-ttu-id="cb1ec-685">説明</span><span class="sxs-lookup"><span data-stu-id="cb1ec-685">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="cb1ec-686">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-686">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="cb1ec-687">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-687">Requirements</span></span>

|<span data-ttu-id="cb1ec-688">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-688">Requirement</span></span>| <span data-ttu-id="cb1ec-689">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-689">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-690">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-690">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-691">1.1</span><span class="sxs-lookup"><span data-stu-id="cb1ec-691">1.1</span></span>|
|[<span data-ttu-id="cb1ec-692">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-692">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-693">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-693">ReadWriteItem</span></span>|
|[<span data-ttu-id="cb1ec-694">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-694">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-695">新規作成</span><span class="sxs-lookup"><span data-stu-id="cb1ec-695">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-696">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-696">Example</span></span>

<span data-ttu-id="cb1ec-697">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-697">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="cb1ec-698">閉じる()</span><span class="sxs-lookup"><span data-stu-id="cb1ec-698">close()</span></span>

<span data-ttu-id="cb1ec-699">新規作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-699">Closes the current item that is being composed.</span></span>

<span data-ttu-id="cb1ec-p137">`close`メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="cb1ec-702">Outlook on the webでは、項目が予定で、`saveAsync`を用いて事前に保存されている場合、項目が最後に保存されてから何も変更されていない場合でも、ユーザーに対して保存、破棄またはキャンセルするよう求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-702">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="cb1ec-703">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close`メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-703">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1ec-704">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-704">Requirements</span></span>

|<span data-ttu-id="cb1ec-705">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-705">Requirement</span></span>| <span data-ttu-id="cb1ec-706">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-706">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-707">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-707">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-708">1.3</span><span class="sxs-lookup"><span data-stu-id="cb1ec-708">1.3</span></span>|
|[<span data-ttu-id="cb1ec-709">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-709">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-710">制限あり</span><span class="sxs-lookup"><span data-stu-id="cb1ec-710">Restricted</span></span>|
|[<span data-ttu-id="cb1ec-711">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-711">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-712">新規作成</span><span class="sxs-lookup"><span data-stu-id="cb1ec-712">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="cb1ec-713">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-713">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="cb1ec-714">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-714">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="cb1ec-715">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-715">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="cb1ec-716">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-716">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="cb1ec-717">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-717">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="cb1ec-p138">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cb1ec-721">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="cb1ec-721">Parameters:</span></span>

| <span data-ttu-id="cb1ec-722">名前</span><span class="sxs-lookup"><span data-stu-id="cb1ec-722">Name</span></span> | <span data-ttu-id="cb1ec-723">種類</span><span class="sxs-lookup"><span data-stu-id="cb1ec-723">Type</span></span> | <span data-ttu-id="cb1ec-724">属性</span><span class="sxs-lookup"><span data-stu-id="cb1ec-724">Attributes</span></span> | <span data-ttu-id="cb1ec-725">説明</span><span class="sxs-lookup"><span data-stu-id="cb1ec-725">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="cb1ec-726">文字列 |オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cb1ec-726">String &#124; Object</span></span>| |<span data-ttu-id="cb1ec-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="cb1ec-729">**または**</span><span class="sxs-lookup"><span data-stu-id="cb1ec-729">**OR**</span></span><br/><span data-ttu-id="cb1ec-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクトです。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="cb1ec-732">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-732">String</span></span> | <span data-ttu-id="cb1ec-733">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-733">&lt;optional&gt;</span></span> | <span data-ttu-id="cb1ec-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="cb1ec-736">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-736">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="cb1ec-737">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-737">&lt;optional&gt;</span></span> | <span data-ttu-id="cb1ec-738">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-738">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="cb1ec-739">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-739">String</span></span> | | <span data-ttu-id="cb1ec-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` でなければいけません。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="cb1ec-742">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-742">String</span></span> | | <span data-ttu-id="cb1ec-743">添付ファイル名を含む文字列です。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-743">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="cb1ec-744">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-744">String</span></span> | | <span data-ttu-id="cb1ec-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URIです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="cb1ec-747">ブール値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-747">Boolean</span></span> | | <span data-ttu-id="cb1ec-p144">`type`が`file`に設定されている場合にのみ使用されます。`true`の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="cb1ec-750">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-750">String</span></span> | | <span data-ttu-id="cb1ec-p145">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムIDです。最大長が 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="cb1ec-754">関数</span><span class="sxs-lookup"><span data-stu-id="cb1ec-754">function</span></span> | <span data-ttu-id="cb1ec-755">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-755">&lt;optional&gt;</span></span> | <span data-ttu-id="cb1ec-756">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである シングル パラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="cb1ec-757">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-757">Requirements</span></span>

|<span data-ttu-id="cb1ec-758">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-758">Requirement</span></span>| <span data-ttu-id="cb1ec-759">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-760">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-760">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-761">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-761">1.0</span></span>|
|[<span data-ttu-id="cb1ec-762">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-762">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-763">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-764">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-764">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-765">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="cb1ec-766">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-766">Examples</span></span>

<span data-ttu-id="cb1ec-767">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-767">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="cb1ec-768">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-768">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="cb1ec-769">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-769">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="cb1ec-770">本文と添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-770">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="cb1ec-771">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-771">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="cb1ec-772">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-772">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="cb1ec-773">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-773">displayReplyForm(formData)</span></span>

<span data-ttu-id="cb1ec-774">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む返信フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-774">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="cb1ec-775">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-775">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="cb1ec-776">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-776">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="cb1ec-777">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-777">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="cb1ec-p146">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cb1ec-781">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="cb1ec-781">Parameters:</span></span>

| <span data-ttu-id="cb1ec-782">名前</span><span class="sxs-lookup"><span data-stu-id="cb1ec-782">Name</span></span> | <span data-ttu-id="cb1ec-783">種類</span><span class="sxs-lookup"><span data-stu-id="cb1ec-783">Type</span></span> | <span data-ttu-id="cb1ec-784">属性</span><span class="sxs-lookup"><span data-stu-id="cb1ec-784">Attributes</span></span> | <span data-ttu-id="cb1ec-785">説明</span><span class="sxs-lookup"><span data-stu-id="cb1ec-785">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="cb1ec-786">文字列 |オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cb1ec-786">String &#124; Object</span></span>| | <span data-ttu-id="cb1ec-p147">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="cb1ec-789">**または**</span><span class="sxs-lookup"><span data-stu-id="cb1ec-789">**OR**</span></span><br/><span data-ttu-id="cb1ec-p148">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクトです。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="cb1ec-792">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-792">String</span></span> | <span data-ttu-id="cb1ec-793">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-793">&lt;optional&gt;</span></span> | <span data-ttu-id="cb1ec-p149">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="cb1ec-796">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-796">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="cb1ec-797">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-797">&lt;optional&gt;</span></span> | <span data-ttu-id="cb1ec-798">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-798">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="cb1ec-799">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-799">String</span></span> | | <span data-ttu-id="cb1ec-p150">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` でなければいけません。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="cb1ec-802">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-802">String</span></span> | | <span data-ttu-id="cb1ec-803">添付ファイル名を含む文字列です。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-803">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="cb1ec-804">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-804">String</span></span> | | <span data-ttu-id="cb1ec-p151">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URIです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="cb1ec-807">ブール値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-807">Boolean</span></span> | | <span data-ttu-id="cb1ec-p152">`type`が`file`に設定されている場合にのみ使用されます。`true`の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="cb1ec-810">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-810">String</span></span> | | <span data-ttu-id="cb1ec-p153">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムIDです。最大長が 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="cb1ec-814">関数</span><span class="sxs-lookup"><span data-stu-id="cb1ec-814">function</span></span> | <span data-ttu-id="cb1ec-815">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-815">&lt;optional&gt;</span></span> | <span data-ttu-id="cb1ec-816">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである シングル パラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-816">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="cb1ec-817">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-817">Requirements</span></span>

|<span data-ttu-id="cb1ec-818">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-818">Requirement</span></span>| <span data-ttu-id="cb1ec-819">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-819">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-820">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-820">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-821">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-821">1.0</span></span>|
|[<span data-ttu-id="cb1ec-822">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-822">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-823">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-823">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-824">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-824">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-825">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-825">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="cb1ec-826">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-826">Examples</span></span>

<span data-ttu-id="cb1ec-827">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-827">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="cb1ec-828">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-828">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="cb1ec-829">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-829">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="cb1ec-830">本文と添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-830">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="cb1ec-831">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-831">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="cb1ec-832">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-832">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="cb1ec-833">getEntities() → {[エンティティ](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="cb1ec-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="cb1ec-834">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-834">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="cb1ec-835">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-835">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1ec-836">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-836">Requirements</span></span>

|<span data-ttu-id="cb1ec-837">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-837">Requirement</span></span>| <span data-ttu-id="cb1ec-838">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-839">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-840">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-840">1.0</span></span>|
|[<span data-ttu-id="cb1ec-841">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-841">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-842">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-843">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-843">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-844">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cb1ec-845">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="cb1ec-845">Returns:</span></span>

<span data-ttu-id="cb1ec-846">種類: [エンティティ](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-846">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="cb1ec-847">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-847">Example</span></span>

<span data-ttu-id="cb1ec-848">次の例では、現在の項目の本文内にある連絡先のエンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-848">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="cb1ec-849">getEntitiesByType(entityType)] → [(空白可能) {<(String|[連絡先](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[電話番号](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion)) >}</span><span class="sxs-lookup"><span data-stu-id="cb1ec-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="cb1ec-850">選択した項目で見つかった指定のエンティティ型のエンティティすべてを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-850">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="cb1ec-851">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-851">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cb1ec-852">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="cb1ec-852">Parameters:</span></span>

|<span data-ttu-id="cb1ec-853">名前</span><span class="sxs-lookup"><span data-stu-id="cb1ec-853">Name</span></span>| <span data-ttu-id="cb1ec-854">種類</span><span class="sxs-lookup"><span data-stu-id="cb1ec-854">Type</span></span>| <span data-ttu-id="cb1ec-855">説明</span><span class="sxs-lookup"><span data-stu-id="cb1ec-855">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="cb1ec-856">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="cb1ec-856">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="cb1ec-857">EntityType 列挙値の 1 つです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-857">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cb1ec-858">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-858">Requirements</span></span>

|<span data-ttu-id="cb1ec-859">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-859">Requirement</span></span>| <span data-ttu-id="cb1ec-860">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-861">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-862">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-862">1.0</span></span>|
|[<span data-ttu-id="cb1ec-863">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-863">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-864">制限あり</span><span class="sxs-lookup"><span data-stu-id="cb1ec-864">Restricted</span></span>|
|[<span data-ttu-id="cb1ec-865">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-865">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-866">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-866">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cb1ec-867">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="cb1ec-867">Returns:</span></span>

<span data-ttu-id="cb1ec-868">`entityType` に渡された値が有効な `EntityType` 列挙型のメンバーでない場合、メソッドは 空白を返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-868">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="cb1ec-869">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-869">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="cb1ec-870">そうでない場合、返される配列内のオブジェクトの種類は、 `entityType` パラメータ内で要求されたエンティティの種類によって異なります。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-870">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="cb1ec-871">このメソッドを使用する最小限のアクセス許可レベルは **制限あり** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-871">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="cb1ec-872">の値 `entityType`</span><span class="sxs-lookup"><span data-stu-id="cb1ec-872">Value of `entityType`</span></span> | <span data-ttu-id="cb1ec-873">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="cb1ec-873">Type of objects in returned array</span></span> | <span data-ttu-id="cb1ec-874">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-874">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="cb1ec-875">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-875">String</span></span> | <span data-ttu-id="cb1ec-876">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="cb1ec-876">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="cb1ec-877">連絡先</span><span class="sxs-lookup"><span data-stu-id="cb1ec-877">Contact</span></span> | <span data-ttu-id="cb1ec-878">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="cb1ec-878">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="cb1ec-879">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-879">String</span></span> | <span data-ttu-id="cb1ec-880">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="cb1ec-880">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="cb1ec-881">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="cb1ec-881">MeetingSuggestion</span></span> | <span data-ttu-id="cb1ec-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="cb1ec-882">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="cb1ec-883">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="cb1ec-883">PhoneNumber</span></span> | <span data-ttu-id="cb1ec-884">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="cb1ec-884">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="cb1ec-885">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="cb1ec-885">TaskSuggestion</span></span> | <span data-ttu-id="cb1ec-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="cb1ec-886">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="cb1ec-887">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-887">String</span></span> | <span data-ttu-id="cb1ec-888">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="cb1ec-888">**Restricted**</span></span> |

<span data-ttu-id="cb1ec-889">型:Array.<(文字列|[連絡先](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="cb1ec-889">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="cb1ec-890">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-890">Example</span></span>

<span data-ttu-id="cb1ec-891">次の例は、現在のアイテムの本文にある郵便アドレスを表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-891">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="cb1ec-892">getFilteredEntitiesByName(name)] → [(Null 許容) {<(文字列| [連絡先](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[電話番号 ](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="cb1ec-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="cb1ec-893">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-893">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="cb1ec-894">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-894">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="cb1ec-895">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-895">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cb1ec-896">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="cb1ec-896">Parameters:</span></span>

|<span data-ttu-id="cb1ec-897">名前</span><span class="sxs-lookup"><span data-stu-id="cb1ec-897">Name</span></span>| <span data-ttu-id="cb1ec-898">種類</span><span class="sxs-lookup"><span data-stu-id="cb1ec-898">Type</span></span>| <span data-ttu-id="cb1ec-899">説明</span><span class="sxs-lookup"><span data-stu-id="cb1ec-899">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="cb1ec-900">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-900">String</span></span>|<span data-ttu-id="cb1ec-901">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-901">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cb1ec-902">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-902">Requirements</span></span>

|<span data-ttu-id="cb1ec-903">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-903">Requirement</span></span>| <span data-ttu-id="cb1ec-904">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-905">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-905">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-906">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-906">1.0</span></span>|
|[<span data-ttu-id="cb1ec-907">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-908">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-909">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-910">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-910">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cb1ec-911">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="cb1ec-911">Returns:</span></span>

<span data-ttu-id="cb1ec-p155">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="cb1ec-914">型:Array.<(文字列|[連絡先](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="cb1ec-914">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="cb1ec-915">getRegExMatches() → {オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-915">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="cb1ec-916">選択した項目内で、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-916">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="cb1ec-917">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-917">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="cb1ec-p156">`getRegExMatches` メソッドは、マニフェスト XML ファイル内にある各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定された項目のプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="cb1ec-921">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-921">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="cb1ec-922">`getRegExMatches` から返されるオブジェクトには、 `fruits` および `veggies` という 2 つのプロパティがありえます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-922">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="cb1ec-p157">アイテムの 本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。代わりに [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb1ec-926">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-926">Requirements</span></span>

|<span data-ttu-id="cb1ec-927">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-927">Requirement</span></span>| <span data-ttu-id="cb1ec-928">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-928">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-929">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-929">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-930">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-930">1.0</span></span>|
|[<span data-ttu-id="cb1ec-931">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-931">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-932">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-932">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-933">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-933">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-934">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-934">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cb1ec-935">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="cb1ec-935">Returns:</span></span>

<span data-ttu-id="cb1ec-p158">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクトです。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="cb1ec-938">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="cb1ec-938">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="cb1ec-939">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cb1ec-939">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="cb1ec-940">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-940">Example</span></span>

<span data-ttu-id="cb1ec-941">次の例は、マニフェストで指定された正規表現の <rule> 要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="cb1ec-941">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="cb1ec-942">getRegExMatchesByName(name)] → [(Null許容) {配列. < 文字列 >}</span><span class="sxs-lookup"><span data-stu-id="cb1ec-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="cb1ec-943">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-943">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="cb1ec-944">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-944">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="cb1ec-945">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-945">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="cb1ec-p159">項目の 本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、項目の本文全体を返さないようにします。`.*` などの正規表現を使用して項目の本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cb1ec-948">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="cb1ec-948">Parameters:</span></span>

|<span data-ttu-id="cb1ec-949">名前</span><span class="sxs-lookup"><span data-stu-id="cb1ec-949">Name</span></span>| <span data-ttu-id="cb1ec-950">種類</span><span class="sxs-lookup"><span data-stu-id="cb1ec-950">Type</span></span>| <span data-ttu-id="cb1ec-951">説明</span><span class="sxs-lookup"><span data-stu-id="cb1ec-951">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="cb1ec-952">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-952">String</span></span>|<span data-ttu-id="cb1ec-953">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-953">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cb1ec-954">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-954">Requirements</span></span>

|<span data-ttu-id="cb1ec-955">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-955">Requirement</span></span>| <span data-ttu-id="cb1ec-956">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-957">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-957">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-958">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-958">1.0</span></span>|
|[<span data-ttu-id="cb1ec-959">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-960">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-960">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-961">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-962">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cb1ec-963">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="cb1ec-963">Returns:</span></span>

<span data-ttu-id="cb1ec-964">マニフェスト XML ファイルで定義された正規表現に一致する文字列を含む配列です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-964">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="cb1ec-965">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="cb1ec-965">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="cb1ec-966">配列. < 文字列 ></span><span class="sxs-lookup"><span data-stu-id="cb1ec-966">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="cb1ec-967">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-967">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="cb1ec-968">getSelectedDataAsync(coercionType, [オプション], コールバック) → {文字列}</span><span class="sxs-lookup"><span data-stu-id="cb1ec-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="cb1ec-969">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-969">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="cb1ec-p160">選択したデータがなく、カーソルが本文または件名にある場合、メソッドは選択したデータに対して Null を返します。本文または件名以外のフィールドが選択されている場合、メソッドは `InvalidSelection` エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cb1ec-972">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="cb1ec-972">Parameters:</span></span>

|<span data-ttu-id="cb1ec-973">名前</span><span class="sxs-lookup"><span data-stu-id="cb1ec-973">Name</span></span>| <span data-ttu-id="cb1ec-974">種類</span><span class="sxs-lookup"><span data-stu-id="cb1ec-974">Type</span></span>| <span data-ttu-id="cb1ec-975">属性</span><span class="sxs-lookup"><span data-stu-id="cb1ec-975">Attributes</span></span>| <span data-ttu-id="cb1ec-976">説明</span><span class="sxs-lookup"><span data-stu-id="cb1ec-976">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="cb1ec-977">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="cb1ec-977">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="cb1ec-p161">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="cb1ec-981">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cb1ec-981">Object</span></span>| <span data-ttu-id="cb1ec-982">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-982">&lt;optional&gt;</span></span>|<span data-ttu-id="cb1ec-983">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-983">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="cb1ec-984">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cb1ec-984">Object</span></span>| <span data-ttu-id="cb1ec-985">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-985">&lt;optional&gt;</span></span>|<span data-ttu-id="cb1ec-986">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-986">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="cb1ec-987">関数</span><span class="sxs-lookup"><span data-stu-id="cb1ec-987">function</span></span>||<span data-ttu-id="cb1ec-988">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-988">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="cb1ec-989">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data`を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-989">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="cb1ec-990">選択元のソース プロパティにアクセスするには、 `asyncResult.value.sourceProperty` を呼び出します。これは `body`   または `subject`  になります。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-990">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cb1ec-991">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-991">Requirements</span></span>

|<span data-ttu-id="cb1ec-992">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-992">Requirement</span></span>| <span data-ttu-id="cb1ec-993">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-993">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-994">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-994">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-995">1.2</span><span class="sxs-lookup"><span data-stu-id="cb1ec-995">1.2</span></span>|
|[<span data-ttu-id="cb1ec-996">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-996">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-997">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-997">ReadWriteItem</span></span>|
|[<span data-ttu-id="cb1ec-998">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-998">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-999">新規作成</span><span class="sxs-lookup"><span data-stu-id="cb1ec-999">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="cb1ec-1000">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1000">Returns:</span></span>

<span data-ttu-id="cb1ec-1001">`coercionType`で決定された書式設定の文字列として選択されたデータです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1001">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="cb1ec-1002">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1002">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="cb1ec-1003">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1003">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="cb1ec-1004">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1004">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="cb1ec-1005">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1005">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="cb1ec-1006">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1006">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="cb1ec-p163">カスタム プロパティは、アプリケーションごと、アイテムごとにキーと値のペアとして保管されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在の項目および現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、項目上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cb1ec-1010">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1010">Parameters:</span></span>

|<span data-ttu-id="cb1ec-1011">名前</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1011">Name</span></span>| <span data-ttu-id="cb1ec-1012">種類</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1012">Type</span></span>| <span data-ttu-id="cb1ec-1013">属性</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1013">Attributes</span></span>| <span data-ttu-id="cb1ec-1014">説明</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1014">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="cb1ec-1015">関数</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1015">function</span></span>||<span data-ttu-id="cb1ec-1016">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="cb1ec-1017">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1017">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="cb1ec-1018">項目からカスタム プロパティを取得、設定、削除して、サーバーにカスタム プロパティのセット バックに対する変更を保存するのに、このオブジェクトが使用できます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1018">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="cb1ec-1019">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1019">Object</span></span>| <span data-ttu-id="cb1ec-1020">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1020">&lt;optional&gt;</span></span>|<span data-ttu-id="cb1ec-1021">開発者は、コールバック 関数でアクセスしたいオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1021">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="cb1ec-1022">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1022">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cb1ec-1023">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1023">Requirements</span></span>

|<span data-ttu-id="cb1ec-1024">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1024">Requirement</span></span>| <span data-ttu-id="cb1ec-1025">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-1026">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1026">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-1027">1.0以降</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1027">1.0</span></span>|
|[<span data-ttu-id="cb1ec-1028">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1028">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1029">ReadItem</span></span>|
|[<span data-ttu-id="cb1ec-1030">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1030">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-1031">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1031">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-1032">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1032">Example</span></span>

<span data-ttu-id="cb1ec-p166">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="cb1ec-1036">removeAttachmentAsync(attachmentId、[オプション]、 [コールバック])</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="cb1ec-1037">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1037">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="cb1ec-p167">`removeAttachmentAsync`メソッドは、指定した識別子の添付ファイルを項目 から削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web アプリ とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別のウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cb1ec-1042">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1042">Parameters:</span></span>

|<span data-ttu-id="cb1ec-1043">名前</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1043">Name</span></span>| <span data-ttu-id="cb1ec-1044">種類</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1044">Type</span></span>| <span data-ttu-id="cb1ec-1045">属性</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1045">Attributes</span></span>| <span data-ttu-id="cb1ec-1046">説明</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1046">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="cb1ec-1047">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1047">String</span></span>||<span data-ttu-id="cb1ec-p168">削除する添付ファイルの識別子です。文字列の最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p168">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="cb1ec-1050">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1050">Object</span></span>| <span data-ttu-id="cb1ec-1051">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1051">&lt;optional&gt;</span></span>|<span data-ttu-id="cb1ec-1052">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1052">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="cb1ec-1053">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1053">Object</span></span>| <span data-ttu-id="cb1ec-1054">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1054">&lt;optional&gt;</span></span>|<span data-ttu-id="cb1ec-1055">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1055">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="cb1ec-1056">関数</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1056">function</span></span>| <span data-ttu-id="cb1ec-1057">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="cb1ec-1058">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="cb1ec-1059">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1059">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="cb1ec-1060">エラー</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1060">Errors</span></span>

| <span data-ttu-id="cb1ec-1061">エラー コード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1061">Error code</span></span> | <span data-ttu-id="cb1ec-1062">説明</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1062">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="cb1ec-1063">添付ファイルの識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1063">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="cb1ec-1064">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1064">Requirements</span></span>

|<span data-ttu-id="cb1ec-1065">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1065">Requirement</span></span>| <span data-ttu-id="cb1ec-1066">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1066">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-1067">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1067">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-1068">1.1</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1068">1.1</span></span>|
|[<span data-ttu-id="cb1ec-1069">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1069">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-1070">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1070">ReadWriteItem</span></span>|
|[<span data-ttu-id="cb1ec-1071">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1071">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-1072">新規作成</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1072">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-1073">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1073">Example</span></span>

<span data-ttu-id="cb1ec-1074">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1074">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="cb1ec-1075">saveAsync ([オプション] 、コールバック)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1075">saveAsync([options], callback)</span></span>

<span data-ttu-id="cb1ec-1076">アイテムを非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1076">Asynchronously saves an item.</span></span>

<span data-ttu-id="cb1ec-p169">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッド経由でアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p169">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="cb1ec-1080">アドインが、WS または REST API を使用しようとして`itemId`を取得するために、新規作成モードでアイテム上の`saveAsync`を呼び出す場合、Outlook キャッシュ モードでは、アイテムがサーバーと実際に同期するまでに時間がかかる場合があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1080">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="cb1ec-1081">項目が同期されるまで、 `itemId` を使用すると、エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1081">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="cb1ec-p171">予定はドラフト状態にはならないため、作成モードで予定に`saveAsync`が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p171">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="cb1ec-1085">次のクライアントは、新規作成モードで予定上の `saveAsync` に対して様々なふるまいをします。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1085">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="cb1ec-1086">Mac Outlook は、作成モードの会議で`saveAsync`をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1086">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="cb1ec-1087">Mac Outlookの会議場で  `saveAsync` を呼びだすと、エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1087">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="cb1ec-1088">作成モードの予定上で`saveAsync`が呼び出されると、Outlook on the web は常に、招待または更新を送信します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1088">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cb1ec-1089">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1089">Parameters:</span></span>

|<span data-ttu-id="cb1ec-1090">名前</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1090">Name</span></span>| <span data-ttu-id="cb1ec-1091">種類</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1091">Type</span></span>| <span data-ttu-id="cb1ec-1092">属性</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1092">Attributes</span></span>| <span data-ttu-id="cb1ec-1093">説明</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1093">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="cb1ec-1094">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1094">Object</span></span>| <span data-ttu-id="cb1ec-1095">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="cb1ec-1096">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1096">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="cb1ec-1097">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1097">Object</span></span>| <span data-ttu-id="cb1ec-1098">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="cb1ec-1099">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1099">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="cb1ec-1100">関数</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1100">function</span></span>||<span data-ttu-id="cb1ec-1101">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="cb1ec-1102">成功すると、アイテム識別子が`asyncResult.value`プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1102">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cb1ec-1103">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1103">Requirements</span></span>

|<span data-ttu-id="cb1ec-1104">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1104">Requirement</span></span>| <span data-ttu-id="cb1ec-1105">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1105">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-1106">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1106">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-1107">1.3</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1107">1.3</span></span>|
|[<span data-ttu-id="cb1ec-1108">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1108">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-1109">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1109">ReadWriteItem</span></span>|
|[<span data-ttu-id="cb1ec-1110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-1111">新規作成</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1111">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="cb1ec-1112">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1112">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="cb1ec-p173">次の例は、コールバック関数に渡される `result` パラメータの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p173">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="cb1ec-1115">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1115">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="cb1ec-1116">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1116">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="cb1ec-p174">`setSelectedDataAsync`メソッドは、指定された文字列を項目のサブジェクトまたは本文のカーソル位置に挿入するか、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p174">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cb1ec-1120">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1120">Parameters:</span></span>

|<span data-ttu-id="cb1ec-1121">名前</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1121">Name</span></span>| <span data-ttu-id="cb1ec-1122">種類</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1122">Type</span></span>| <span data-ttu-id="cb1ec-1123">属性</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1123">Attributes</span></span>| <span data-ttu-id="cb1ec-1124">説明</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1124">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="cb1ec-1125">文字列</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1125">String</span></span>||<span data-ttu-id="cb1ec-p175">挿入されるデータです。データの長さは 1,000,000 文字以内です。1,000,000 文字を超えるデータが渡されると、 `ArgumentOutOfRange` の例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p175">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="cb1ec-1129">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1129">Object</span></span>| <span data-ttu-id="cb1ec-1130">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="cb1ec-1131">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1131">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="cb1ec-1132">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1132">Object</span></span>| <span data-ttu-id="cb1ec-1133">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="cb1ec-1134">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1134">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="cb1ec-1135">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1135">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="cb1ec-1136">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1136">&lt;optional&gt;</span></span>|<span data-ttu-id="cb1ec-p176">`text` の場合、Outlook Web アプリ と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p176">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="cb1ec-p177">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-p177">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="cb1ec-1141">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1141">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="cb1ec-1142">関数</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1142">function</span></span>||<span data-ttu-id="cb1ec-1143">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1143">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="cb1ec-1144">要件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1144">Requirements</span></span>

|<span data-ttu-id="cb1ec-1145">必要条件</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1145">Requirement</span></span>| <span data-ttu-id="cb1ec-1146">値</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1146">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb1ec-1147">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1147">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb1ec-1148">1.2</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1148">1.2</span></span>|
|[<span data-ttu-id="cb1ec-1149">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1149">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb1ec-1150">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1150">ReadWriteItem</span></span>|
|[<span data-ttu-id="cb1ec-1151">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cb1ec-1152">新規作成</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1152">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="cb1ec-1153">例</span><span class="sxs-lookup"><span data-stu-id="cb1ec-1153">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```