
# <a name="item"></a><span data-ttu-id="c3b6b-101">アイテム</span><span class="sxs-lookup"><span data-stu-id="c3b6b-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="c3b6b-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="c3b6b-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="c3b6b-p101">`item` 名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。`item` の種類を [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) プロパティを使用して指定できます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-105">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-105">Requirements</span></span>

|<span data-ttu-id="c3b6b-106">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-106">Requirement</span></span>| <span data-ttu-id="c3b6b-107">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-108">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-109">1.0</span></span>|
|[<span data-ttu-id="c3b6b-110">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="c3b6b-111">Restricted</span></span>|
|[<span data-ttu-id="c3b6b-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-113">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c3b6b-114">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="c3b6b-114">Members and methods</span></span>

| <span data-ttu-id="c3b6b-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-115">Member</span></span> | <span data-ttu-id="c3b6b-116">型</span><span class="sxs-lookup"><span data-stu-id="c3b6b-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c3b6b-117">添付ファイル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails) | <span data-ttu-id="c3b6b-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-118">Member</span></span> |
| [<span data-ttu-id="c3b6b-119">BCC</span><span class="sxs-lookup"><span data-stu-id="c3b6b-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="c3b6b-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-120">Member</span></span> |
| [<span data-ttu-id="c3b6b-121">本文</span><span class="sxs-lookup"><span data-stu-id="c3b6b-121">body</span></span>](#body-bodyjavascriptapioutlook16officebody) | <span data-ttu-id="c3b6b-122">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-122">Member</span></span> |
| [<span data-ttu-id="c3b6b-123">CC</span><span class="sxs-lookup"><span data-stu-id="c3b6b-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="c3b6b-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-124">Member</span></span> |
| [<span data-ttu-id="c3b6b-125">会話 ID</span><span class="sxs-lookup"><span data-stu-id="c3b6b-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="c3b6b-126">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-126">Member</span></span> |
| [<span data-ttu-id="c3b6b-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="c3b6b-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="c3b6b-128">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-128">Member</span></span> |
| [<span data-ttu-id="c3b6b-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="c3b6b-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="c3b6b-130">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-130">Member</span></span> |
| [<span data-ttu-id="c3b6b-131">終了</span><span class="sxs-lookup"><span data-stu-id="c3b6b-131">end</span></span>](#end-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="c3b6b-132">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-132">Member</span></span> |
| [<span data-ttu-id="c3b6b-133">送信者</span><span class="sxs-lookup"><span data-stu-id="c3b6b-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="c3b6b-134">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-134">Member</span></span> |
| [<span data-ttu-id="c3b6b-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="c3b6b-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="c3b6b-136">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-136">Member</span></span> |
| [<span data-ttu-id="c3b6b-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="c3b6b-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="c3b6b-138">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-138">Member</span></span> |
| [<span data-ttu-id="c3b6b-139">itemId</span><span class="sxs-lookup"><span data-stu-id="c3b6b-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="c3b6b-140">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-140">Member</span></span> |
| [<span data-ttu-id="c3b6b-141">itemType</span><span class="sxs-lookup"><span data-stu-id="c3b6b-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) | <span data-ttu-id="c3b6b-142">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-142">Member</span></span> |
| [<span data-ttu-id="c3b6b-143">位置</span><span class="sxs-lookup"><span data-stu-id="c3b6b-143">location</span></span>](#location-stringlocationjavascriptapioutlook16officelocation) | <span data-ttu-id="c3b6b-144">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-144">Member</span></span> |
| [<span data-ttu-id="c3b6b-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="c3b6b-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="c3b6b-146">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-146">Member</span></span> |
| [<span data-ttu-id="c3b6b-147">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="c3b6b-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages) | <span data-ttu-id="c3b6b-148">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-148">Member</span></span> |
| [<span data-ttu-id="c3b6b-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="c3b6b-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="c3b6b-150">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-150">Member</span></span> |
| [<span data-ttu-id="c3b6b-151">主催者</span><span class="sxs-lookup"><span data-stu-id="c3b6b-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="c3b6b-152">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-152">Member</span></span> |
| [<span data-ttu-id="c3b6b-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="c3b6b-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="c3b6b-154">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-154">Member</span></span> |
| [<span data-ttu-id="c3b6b-155">差出人</span><span class="sxs-lookup"><span data-stu-id="c3b6b-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="c3b6b-156">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-156">Member</span></span> |
| [<span data-ttu-id="c3b6b-157">開始</span><span class="sxs-lookup"><span data-stu-id="c3b6b-157">start</span></span>](#start-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="c3b6b-158">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-158">Member</span></span> |
| [<span data-ttu-id="c3b6b-159">件名</span><span class="sxs-lookup"><span data-stu-id="c3b6b-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook16officesubject) | <span data-ttu-id="c3b6b-160">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-160">Member</span></span> |
| [<span data-ttu-id="c3b6b-161">宛先</span><span class="sxs-lookup"><span data-stu-id="c3b6b-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="c3b6b-162">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-162">Member</span></span> |
| [<span data-ttu-id="c3b6b-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c3b6b-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="c3b6b-164">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3b6b-164">Method</span></span> |
| [<span data-ttu-id="c3b6b-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c3b6b-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="c3b6b-166">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3b6b-166">Method</span></span> |
| [<span data-ttu-id="c3b6b-167">終了</span><span class="sxs-lookup"><span data-stu-id="c3b6b-167">close</span></span>](#close) | <span data-ttu-id="c3b6b-168">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3b6b-168">Method</span></span> |
| [<span data-ttu-id="c3b6b-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="c3b6b-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="c3b6b-170">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3b6b-170">Method</span></span> |
| [<span data-ttu-id="c3b6b-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="c3b6b-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="c3b6b-172">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3b6b-172">Method</span></span> |
| [<span data-ttu-id="c3b6b-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="c3b6b-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="c3b6b-174">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3b6b-174">Method</span></span> |
| [<span data-ttu-id="c3b6b-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="c3b6b-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="c3b6b-176">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3b6b-176">Method</span></span> |
| [<span data-ttu-id="c3b6b-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="c3b6b-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="c3b6b-178">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3b6b-178">Method</span></span> |
| [<span data-ttu-id="c3b6b-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="c3b6b-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="c3b6b-180">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3b6b-180">Method</span></span> |
| [<span data-ttu-id="c3b6b-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="c3b6b-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="c3b6b-182">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3b6b-182">Method</span></span> |
| [<span data-ttu-id="c3b6b-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c3b6b-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="c3b6b-184">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3b6b-184">Method</span></span> |
| [<span data-ttu-id="c3b6b-185">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="c3b6b-185">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="c3b6b-186">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3b6b-186">Method</span></span> |
| [<span data-ttu-id="c3b6b-187">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="c3b6b-187">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="c3b6b-188">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3b6b-188">Method</span></span> |
| [<span data-ttu-id="c3b6b-189">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="c3b6b-189">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="c3b6b-190">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3b6b-190">Method</span></span> |
| [<span data-ttu-id="c3b6b-191">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c3b6b-191">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="c3b6b-192">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3b6b-192">Method</span></span> |
| [<span data-ttu-id="c3b6b-193">saveAsync</span><span class="sxs-lookup"><span data-stu-id="c3b6b-193">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="c3b6b-194">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3b6b-194">Method</span></span> |
| [<span data-ttu-id="c3b6b-195">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c3b6b-195">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="c3b6b-196">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3b6b-196">Method</span></span> |

### <a name="example"></a><span data-ttu-id="c3b6b-197">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-197">Example</span></span>

<span data-ttu-id="c3b6b-198">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-198">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="c3b6b-199">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-199">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="c3b6b-200">添付ファイル :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c3b6b-200">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="c3b6b-p102">アイテムの添付ファイルの配列を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c3b6b-203">潜在的なセキュリティ問題により、特定の種類のファイルは Outlook でブロックされ、したがって戻ってきません。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-203">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="c3b6b-204">詳細については、 [Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-204">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-205">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-205">Type:</span></span>

*   <span data-ttu-id="c3b6b-206">配列.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c3b6b-206">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-207">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-207">Requirements</span></span>

|<span data-ttu-id="c3b6b-208">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-208">Requirement</span></span>| <span data-ttu-id="c3b6b-209">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-210">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-211">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-211">1.0</span></span>|
|[<span data-ttu-id="c3b6b-212">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-213">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-214">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-215">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-216">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-216">Example</span></span>

<span data-ttu-id="c3b6b-217">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-217">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="c3b6b-218">BCC:[受信者](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-218">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="c3b6b-219">メッセージの bcc (ブラインド カーボン コピー) 列上の 受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-219">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="c3b6b-220">作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-220">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-221">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-221">Type:</span></span>

*   [<span data-ttu-id="c3b6b-222">受信者</span><span class="sxs-lookup"><span data-stu-id="c3b6b-222">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="c3b6b-223">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-223">Requirements</span></span>

|<span data-ttu-id="c3b6b-224">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-224">Requirement</span></span>| <span data-ttu-id="c3b6b-225">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-226">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-226">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-227">1.1</span><span class="sxs-lookup"><span data-stu-id="c3b6b-227">1.1</span></span>|
|[<span data-ttu-id="c3b6b-228">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-229">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-230">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-231">作成</span><span class="sxs-lookup"><span data-stu-id="c3b6b-231">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-232">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-232">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="c3b6b-233">本文:[本文](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-233">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="c3b6b-234">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-234">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-235">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-235">Type:</span></span>

*   [<span data-ttu-id="c3b6b-236">本文</span><span class="sxs-lookup"><span data-stu-id="c3b6b-236">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="c3b6b-237">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-237">Requirements</span></span>

|<span data-ttu-id="c3b6b-238">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-238">Requirement</span></span>| <span data-ttu-id="c3b6b-239">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-240">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-241">1.1</span><span class="sxs-lookup"><span data-stu-id="c3b6b-241">1.1</span></span>|
|[<span data-ttu-id="c3b6b-242">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-243">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-244">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-245">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-245">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="c3b6b-246">cc: 配列. <[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="c3b6b-247">メッセージの Cc (カーボン コピー) 受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="c3b6b-248">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3b6b-249">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-249">Read mode</span></span>

<span data-ttu-id="c3b6b-p106">`cc` プロパティは、 `EmailAddressDetails` オブ ジェクトを含む配列をメッセージの **Cc** 行にある各受信者について返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c3b6b-252">作成モード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-252">Compose mode</span></span>

<span data-ttu-id="c3b6b-253">`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-253">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-254">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-254">Type:</span></span>

*   <span data-ttu-id="c3b6b-255">配列.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-256">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-256">Requirements</span></span>

|<span data-ttu-id="c3b6b-257">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-257">Requirement</span></span>| <span data-ttu-id="c3b6b-258">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-259">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-259">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-260">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-260">1.0</span></span>|
|[<span data-ttu-id="c3b6b-261">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-261">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-262">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-263">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-264">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-264">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-265">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-265">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="c3b6b-266">（Null 許容）conversationId：文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-266">(nullable) conversationId :String</span></span>

<span data-ttu-id="c3b6b-267">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-267">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="c3b6b-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="c3b6b-p108">作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定しアイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-272">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-272">Type:</span></span>

*   <span data-ttu-id="c3b6b-273">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-273">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-274">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-274">Requirements</span></span>

|<span data-ttu-id="c3b6b-275">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-275">Requirement</span></span>| <span data-ttu-id="c3b6b-276">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-277">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-277">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-278">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-278">1.0</span></span>|
|[<span data-ttu-id="c3b6b-279">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-279">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-280">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-281">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-281">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-282">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-282">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="c3b6b-283">dateTimeCreated: 日付</span><span class="sxs-lookup"><span data-stu-id="c3b6b-283">dateTimeCreated :Date</span></span>

<span data-ttu-id="c3b6b-p109">アイテムが作成された日時を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-286">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-286">Type:</span></span>

*   <span data-ttu-id="c3b6b-287">日付</span><span class="sxs-lookup"><span data-stu-id="c3b6b-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-288">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-288">Requirements</span></span>

|<span data-ttu-id="c3b6b-289">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-289">Requirement</span></span>| <span data-ttu-id="c3b6b-290">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-291">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-291">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-292">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-292">1.0</span></span>|
|[<span data-ttu-id="c3b6b-293">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-294">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-295">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-296">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-297">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-297">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="c3b6b-298">dateTimeModified: 日付</span><span class="sxs-lookup"><span data-stu-id="c3b6b-298">dateTimeModified :Date</span></span>

<span data-ttu-id="c3b6b-p110">アイテムが最後に変更された日時を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c3b6b-301">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-301">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-302">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-302">Type:</span></span>

*   <span data-ttu-id="c3b6b-303">日付</span><span class="sxs-lookup"><span data-stu-id="c3b6b-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-304">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-304">Requirements</span></span>

|<span data-ttu-id="c3b6b-305">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-305">Requirement</span></span>| <span data-ttu-id="c3b6b-306">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-307">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-307">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-308">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-308">1.0</span></span>|
|[<span data-ttu-id="c3b6b-309">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-309">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-310">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-311">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-311">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-312">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-313">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-313">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="c3b6b-314">end :日付 |[時間](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-314">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="c3b6b-315">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="c3b6b-p111">`end` プロパティは、協定世界時 (UTC) 形式の時刻値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3b6b-318">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-318">Read mode</span></span>

<span data-ttu-id="c3b6b-319">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-319">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c3b6b-320">作成モード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-320">Compose mode</span></span>

<span data-ttu-id="c3b6b-321">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="c3b6b-322">[ `Time.setAsync` ](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-322">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-323">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-323">Type:</span></span>

*   <span data-ttu-id="c3b6b-324">日付| [時間](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-324">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-325">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-325">Requirements</span></span>

|<span data-ttu-id="c3b6b-326">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-326">Requirement</span></span>| <span data-ttu-id="c3b6b-327">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-328">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-328">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-329">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-329">1.0</span></span>|
|[<span data-ttu-id="c3b6b-330">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-330">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-331">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-332">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-332">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-333">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-333">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-334">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-334">Example</span></span>

<span data-ttu-id="c3b6b-335">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して、作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-335">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="c3b6b-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="c3b6b-p112">メッセージの送信者の電子メール アドレスを取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="c3b6b-p113">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) プロパティは同一人物を表します。その場合、`from` プロパティは委任者を、送信者プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c3b6b-341">`from` プロパティ内の`EmailAddressDetails` オブジェクトの`recipientType`  プロパティは、 `undefined`です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-341">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-342">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-342">Type:</span></span>

*   [<span data-ttu-id="c3b6b-343">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c3b6b-343">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="c3b6b-344">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-344">Requirements</span></span>

|<span data-ttu-id="c3b6b-345">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-345">Requirement</span></span>| <span data-ttu-id="c3b6b-346">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-346">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-347">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-347">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-348">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-348">1.0</span></span>|
|[<span data-ttu-id="c3b6b-349">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-349">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-350">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-350">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-351">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-351">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-352">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-352">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="c3b6b-353">internetMessageId: 文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-353">internetMessageId :String</span></span>

<span data-ttu-id="c3b6b-p114">電子メール メッセージのインターネット メッセージの識別子を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-356">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-356">Type:</span></span>

*   <span data-ttu-id="c3b6b-357">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-358">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-358">Requirements</span></span>

|<span data-ttu-id="c3b6b-359">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-359">Requirement</span></span>| <span data-ttu-id="c3b6b-360">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-361">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-361">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-362">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-362">1.0</span></span>|
|[<span data-ttu-id="c3b6b-363">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-364">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-365">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-366">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-367">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-367">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="c3b6b-368">itemClass: 文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-368">itemClass :String</span></span>

<span data-ttu-id="c3b6b-p115">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="c3b6b-p116">`itemClass` プロパティは、選択したアイテムのメッセージ クラスを指定します。次は、メッセージまたは予定アイテムの既定のメッセージ クラスを示しています。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="c3b6b-373">型</span><span class="sxs-lookup"><span data-stu-id="c3b6b-373">Type</span></span> | <span data-ttu-id="c3b6b-374">説明</span><span class="sxs-lookup"><span data-stu-id="c3b6b-374">Description</span></span> | <span data-ttu-id="c3b6b-375">アイテムクラス</span><span class="sxs-lookup"><span data-stu-id="c3b6b-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="c3b6b-376">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="c3b6b-376">Appointment items</span></span> | <span data-ttu-id="c3b6b-377">これらは、アイテムクラス `IPM.Appointment` または `IPM.Appointment.Occurence` の予定表アイテムです。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="c3b6b-378">メッセージアイテム</span><span class="sxs-lookup"><span data-stu-id="c3b6b-378">Message items</span></span> | <span data-ttu-id="c3b6b-379">これには、基本のメッセージ クラス として `IPM.Schedule.Meeting`  を使用する、既定のメッセージ クラス `IPM.Note`  会議出席依頼、および取り消しを持つ電子メール メッセージが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="c3b6b-380">既定のメッセージ クラスを拡張したカスタム メッセージ クラス 、たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` などを作成できます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-381">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-381">Type:</span></span>

*   <span data-ttu-id="c3b6b-382">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-383">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-383">Requirements</span></span>

|<span data-ttu-id="c3b6b-384">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-384">Requirement</span></span>| <span data-ttu-id="c3b6b-385">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-386">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-386">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-387">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-387">1.0</span></span>|
|[<span data-ttu-id="c3b6b-388">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-388">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-389">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-390">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-390">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-391">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-392">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-392">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="c3b6b-393">（Null 許容） itemId ：文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-393">(nullable) itemId :String</span></span>

<span data-ttu-id="c3b6b-p117">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c3b6b-396">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="c3b6b-397">`itemId` プロパティは、Outlook の Entry ID または Outlook の REST API によって使用される ID と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="c3b6b-398">この値を使用して REST API 呼び出しを行う前に [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="c3b6b-399">詳細については、 [Outlook アドインから Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="c3b6b-p119">作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメータでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-402">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-402">Type:</span></span>

*   <span data-ttu-id="c3b6b-403">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-404">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-404">Requirements</span></span>

|<span data-ttu-id="c3b6b-405">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-405">Requirement</span></span>| <span data-ttu-id="c3b6b-406">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-407">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-407">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-408">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-408">1.0</span></span>|
|[<span data-ttu-id="c3b6b-409">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-409">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-410">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-411">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-411">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-412">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-413">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-413">Example</span></span>

<span data-ttu-id="c3b6b-p120">次のコードは、アイテム識別子のプレゼンスを確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="c3b6b-416">itemType:[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="c3b6b-417">インスタンスが表しているアイテムの型を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="c3b6b-418">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-419">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-419">Type:</span></span>

*   [<span data-ttu-id="c3b6b-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="c3b6b-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="c3b6b-421">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-421">Requirements</span></span>

|<span data-ttu-id="c3b6b-422">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-422">Requirement</span></span>| <span data-ttu-id="c3b6b-423">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-424">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-424">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-425">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-425">1.0</span></span>|
|[<span data-ttu-id="c3b6b-426">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-427">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-428">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-429">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-429">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-430">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-430">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="c3b6b-431">場所: 文字列|[場所](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-431">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="c3b6b-432">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3b6b-433">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-433">Read mode</span></span>

<span data-ttu-id="c3b6b-434">`location` プロパティは、予定の場所を含む文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-434">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c3b6b-435">作成モード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-435">Compose mode</span></span>

<span data-ttu-id="c3b6b-436">`location` プロパティは、予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-437">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-437">Type:</span></span>

*   <span data-ttu-id="c3b6b-438">文字列 | [場所](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-438">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-439">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-439">Requirements</span></span>

|<span data-ttu-id="c3b6b-440">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-440">Requirement</span></span>| <span data-ttu-id="c3b6b-441">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-442">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-442">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-443">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-443">1.0</span></span>|
|[<span data-ttu-id="c3b6b-444">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-445">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-446">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-447">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-448">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-448">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="c3b6b-449">normalizedSubject :文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-449">normalizedSubject :String</span></span>

<span data-ttu-id="c3b6b-p121">すべてのプレフィックス (`RE:` や `FWD:` を含む) が削除されたアイテムの件名を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="c3b6b-p122">normalizedSubject プロパティは、電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたアイテムの件名を取得します。プレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-454">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-454">Type:</span></span>

*   <span data-ttu-id="c3b6b-455">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-455">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-456">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-456">Requirements</span></span>

|<span data-ttu-id="c3b6b-457">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-457">Requirement</span></span>| <span data-ttu-id="c3b6b-458">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-458">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-459">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-459">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-460">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-460">1.0</span></span>|
|[<span data-ttu-id="c3b6b-461">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-461">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-462">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-462">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-463">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-463">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-464">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-464">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-465">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-465">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="c3b6b-466">notificationMessages:[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-466">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="c3b6b-467">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-467">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-468">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-468">Type:</span></span>

*   [<span data-ttu-id="c3b6b-469">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="c3b6b-469">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="c3b6b-470">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-470">Requirements</span></span>

|<span data-ttu-id="c3b6b-471">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-471">Requirement</span></span>| <span data-ttu-id="c3b6b-472">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-472">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-473">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-473">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-474">1.3</span><span class="sxs-lookup"><span data-stu-id="c3b6b-474">1.3</span></span>|
|[<span data-ttu-id="c3b6b-475">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-475">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-476">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-476">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-477">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-477">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-478">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-478">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="c3b6b-479">optionalAttendees: 配列.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="c3b6b-480">イベントの任意の出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="c3b6b-481">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3b6b-482">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-482">Read mode</span></span>

<span data-ttu-id="c3b6b-483">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c3b6b-484">作成モード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-484">Compose mode</span></span>

<span data-ttu-id="c3b6b-485">`optionalAttendees` プロパティは会議への任意出席者を取得および設定するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-486">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-486">Type:</span></span>

*   <span data-ttu-id="c3b6b-487">配列.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> |[受信者](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-488">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-488">Requirements</span></span>

|<span data-ttu-id="c3b6b-489">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-489">Requirement</span></span>| <span data-ttu-id="c3b6b-490">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-491">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-491">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-492">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-492">1.0</span></span>|
|[<span data-ttu-id="c3b6b-493">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-493">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-494">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-495">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-495">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-496">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-496">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-497">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-497">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="c3b6b-498">オーガナイザー:[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-498">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="c3b6b-p124">指定の会議の開催者の電子メールアドレスを取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-501">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-501">Type:</span></span>

*   [<span data-ttu-id="c3b6b-502">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c3b6b-502">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="c3b6b-503">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-503">Requirements</span></span>

|<span data-ttu-id="c3b6b-504">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-504">Requirement</span></span>| <span data-ttu-id="c3b6b-505">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-506">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-506">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-507">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-507">1.0</span></span>|
|[<span data-ttu-id="c3b6b-508">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-509">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-510">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-511">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-511">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-512">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-512">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="c3b6b-513">requiredAttendees: 配列.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-513">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="c3b6b-514">イベントの必須の出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-514">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="c3b6b-515">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-515">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3b6b-516">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-516">Read mode</span></span>

<span data-ttu-id="c3b6b-517">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-517">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c3b6b-518">作成モード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-518">Compose mode</span></span>

<span data-ttu-id="c3b6b-519">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-519">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-520">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-520">Type:</span></span>

*   <span data-ttu-id="c3b6b-521">配列.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-521">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-522">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-522">Requirements</span></span>

|<span data-ttu-id="c3b6b-523">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-523">Requirement</span></span>| <span data-ttu-id="c3b6b-524">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-525">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-525">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-526">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-526">1.0</span></span>|
|[<span data-ttu-id="c3b6b-527">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-527">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-528">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-529">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-529">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-530">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-530">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-531">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-531">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="c3b6b-532">送信者:[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-532">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="c3b6b-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="c3b6b-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。その場合、`from` プロパティは委任者を、送信者プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c3b6b-537">`sender` プロパティ内の`EmailAddressDetails` オブジェクトの`recipientType`  プロパティは、 `undefined`です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-537">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-538">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-538">Type:</span></span>

*   [<span data-ttu-id="c3b6b-539">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c3b6b-539">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="c3b6b-540">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-540">Requirements</span></span>

|<span data-ttu-id="c3b6b-541">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-541">Requirement</span></span>| <span data-ttu-id="c3b6b-542">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-542">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-543">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-543">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-544">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-544">1.0</span></span>|
|[<span data-ttu-id="c3b6b-545">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-545">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-546">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-546">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-547">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-547">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-548">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-548">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-549">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-549">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="c3b6b-550">開始: 日付 |[時間](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-550">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="c3b6b-551">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-551">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="c3b6b-p128">`start` プロパティは、協定世界時 (UTC) 形式の時刻値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3b6b-554">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-554">Read mode</span></span>

<span data-ttu-id="c3b6b-555">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-555">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c3b6b-556">作成モード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-556">Compose mode</span></span>

<span data-ttu-id="c3b6b-557">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-557">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="c3b6b-558">[ `Time.setAsync` ](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-558">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-559">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-559">Type:</span></span>

*   <span data-ttu-id="c3b6b-560">日付| [時間](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-560">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-561">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-561">Requirements</span></span>

|<span data-ttu-id="c3b6b-562">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-562">Requirement</span></span>| <span data-ttu-id="c3b6b-563">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-564">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-564">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-565">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-565">1.0</span></span>|
|[<span data-ttu-id="c3b6b-566">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-567">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-568">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-569">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-569">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-570">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-570">Example</span></span>

<span data-ttu-id="c3b6b-571">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-571">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="c3b6b-572">件名: 文字列|[件名](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-572">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="c3b6b-573">アイテムの件名フィールドに表示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-573">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="c3b6b-574">`subject` プロパティは、電子メールサーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-574">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3b6b-575">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-575">Read mode</span></span>

<span data-ttu-id="c3b6b-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような行間のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="c3b6b-578">作成モード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-578">Compose mode</span></span>

<span data-ttu-id="c3b6b-579">`subject` プロパティは、件名を取得または設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-579">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c3b6b-580">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-580">Type:</span></span>

*   <span data-ttu-id="c3b6b-581">文字列 | [件名](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-581">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-582">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-582">Requirements</span></span>

|<span data-ttu-id="c3b6b-583">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-583">Requirement</span></span>| <span data-ttu-id="c3b6b-584">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-585">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-585">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-586">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-586">1.0</span></span>|
|[<span data-ttu-id="c3b6b-587">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-587">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-588">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-588">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-589">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-589">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-590">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-590">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="c3b6b-591">to :配列.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-591">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="c3b6b-592">メッセージの  **宛先**行の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-592">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="c3b6b-593">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-593">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3b6b-594">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-594">Read mode</span></span>

<span data-ttu-id="c3b6b-p131">`to` プロパティは、   `EmailAddressDetails` オブジェクトを含む配列を、メッセージの  **To** 行にある各受信者について、 返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c3b6b-597">作成モード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-597">Compose mode</span></span>

<span data-ttu-id="c3b6b-598">`to` プロパティは、メッセージの **To** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-598">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="c3b6b-599">型:</span><span class="sxs-lookup"><span data-stu-id="c3b6b-599">Type:</span></span>

*   <span data-ttu-id="c3b6b-600">配列.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-600">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-601">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-601">Requirements</span></span>

|<span data-ttu-id="c3b6b-602">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-602">Requirement</span></span>| <span data-ttu-id="c3b6b-603">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-603">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-604">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-604">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-605">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-605">1.0</span></span>|
|[<span data-ttu-id="c3b6b-606">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-606">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-607">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-607">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-608">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-608">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-609">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-609">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-610">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-610">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="c3b6b-611">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3b6b-611">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="c3b6b-612">addFileAttachmentAsync (uri、attachmentName、[オプション]、[コールバック])</span><span class="sxs-lookup"><span data-stu-id="c3b6b-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c3b6b-613">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="c3b6b-614">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="c3b6b-615">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3b6b-616">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="c3b6b-616">Parameters:</span></span>

|<span data-ttu-id="c3b6b-617">名前</span><span class="sxs-lookup"><span data-stu-id="c3b6b-617">Name</span></span>| <span data-ttu-id="c3b6b-618">型</span><span class="sxs-lookup"><span data-stu-id="c3b6b-618">Type</span></span>| <span data-ttu-id="c3b6b-619">属性</span><span class="sxs-lookup"><span data-stu-id="c3b6b-619">Attributes</span></span>| <span data-ttu-id="c3b6b-620">説明</span><span class="sxs-lookup"><span data-stu-id="c3b6b-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="c3b6b-621">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-621">String</span></span>||<span data-ttu-id="c3b6b-p132">メッセージまたは予定に添付するファイルの場所を示す URIです。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="c3b6b-624">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-624">String</span></span>||<span data-ttu-id="c3b6b-p133">添付ファイルのアップロード時に表示される添付ファイルの名前です。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="c3b6b-627">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3b6b-627">Object</span></span>| <span data-ttu-id="c3b6b-628">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-628">&lt;optional&gt;</span></span>|<span data-ttu-id="c3b6b-629">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="c3b6b-630">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3b6b-630">Object</span></span> | <span data-ttu-id="c3b6b-631">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-631">&lt;optional&gt;</span></span> | <span data-ttu-id="c3b6b-632">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="c3b6b-633">ブール値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-633">Boolean</span></span> | <span data-ttu-id="c3b6b-634">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-634">&lt;optional&gt;</span></span> | <span data-ttu-id="c3b6b-635">`true`の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="c3b6b-636">関数</span><span class="sxs-lookup"><span data-stu-id="c3b6b-636">function</span></span>| <span data-ttu-id="c3b6b-637">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-637">&lt;optional&gt;</span></span>|<span data-ttu-id="c3b6b-638">メソッドが完了すると、`callback`パラメータに渡された関数が、シングル パラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c3b6b-639">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c3b6b-640">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c3b6b-641">エラー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-641">Errors</span></span>

| <span data-ttu-id="c3b6b-642">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-642">Error code</span></span> | <span data-ttu-id="c3b6b-643">説明</span><span class="sxs-lookup"><span data-stu-id="c3b6b-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="c3b6b-644">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="c3b6b-645">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="c3b6b-646">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c3b6b-647">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-647">Requirements</span></span>

|<span data-ttu-id="c3b6b-648">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-648">Requirement</span></span>| <span data-ttu-id="c3b6b-649">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-650">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-650">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-651">1.1</span><span class="sxs-lookup"><span data-stu-id="c3b6b-651">1.1</span></span>|
|[<span data-ttu-id="c3b6b-652">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3b6b-654">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-655">作成</span><span class="sxs-lookup"><span data-stu-id="c3b6b-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c3b6b-656">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-656">Examples</span></span>

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

<span data-ttu-id="c3b6b-657">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="c3b6b-658">addItemAttachmentAsync (itemd、attachmentName、[オプション]、[コールバック])</span><span class="sxs-lookup"><span data-stu-id="c3b6b-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c3b6b-659">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="c3b6b-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメータがあるメソッドが呼び出されます。このパラメータには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメータを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="c3b6b-663">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="c3b6b-664">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、推奨されていません。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-664">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3b6b-665">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="c3b6b-665">Parameters:</span></span>

|<span data-ttu-id="c3b6b-666">名前</span><span class="sxs-lookup"><span data-stu-id="c3b6b-666">Name</span></span>| <span data-ttu-id="c3b6b-667">型</span><span class="sxs-lookup"><span data-stu-id="c3b6b-667">Type</span></span>| <span data-ttu-id="c3b6b-668">属性</span><span class="sxs-lookup"><span data-stu-id="c3b6b-668">Attributes</span></span>| <span data-ttu-id="c3b6b-669">説明</span><span class="sxs-lookup"><span data-stu-id="c3b6b-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="c3b6b-670">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-670">String</span></span>||<span data-ttu-id="c3b6b-p135">添付するアイテムの Exchange 識別子です。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="c3b6b-673">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-673">String</span></span>||<span data-ttu-id="c3b6b-p136">添付するアイテムの件名です。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="c3b6b-676">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3b6b-676">Object</span></span>| <span data-ttu-id="c3b6b-677">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-677">&lt;optional&gt;</span></span>|<span data-ttu-id="c3b6b-678">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c3b6b-679">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3b6b-679">Object</span></span>| <span data-ttu-id="c3b6b-680">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-680">&lt;optional&gt;</span></span>|<span data-ttu-id="c3b6b-681">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c3b6b-682">関数</span><span class="sxs-lookup"><span data-stu-id="c3b6b-682">function</span></span>| <span data-ttu-id="c3b6b-683">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-683">&lt;optional&gt;</span></span>|<span data-ttu-id="c3b6b-684">メソッドが完了すると、`callback`パラメータに渡された関数が、シングル パラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c3b6b-685">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c3b6b-686">添付ファイルの追加に失敗した場合、 エラーの説明を提供する`asyncResult` オブジェクトが `Error` オブジェクトに含まれます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c3b6b-687">エラー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-687">Errors</span></span>

| <span data-ttu-id="c3b6b-688">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-688">Error code</span></span> | <span data-ttu-id="c3b6b-689">説明</span><span class="sxs-lookup"><span data-stu-id="c3b6b-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="c3b6b-690">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c3b6b-691">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-691">Requirements</span></span>

|<span data-ttu-id="c3b6b-692">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-692">Requirement</span></span>| <span data-ttu-id="c3b6b-693">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-694">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-694">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-695">1.1</span><span class="sxs-lookup"><span data-stu-id="c3b6b-695">1.1</span></span>|
|[<span data-ttu-id="c3b6b-696">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-696">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3b6b-698">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-698">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-699">作成</span><span class="sxs-lookup"><span data-stu-id="c3b6b-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-700">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-700">Example</span></span>

<span data-ttu-id="c3b6b-701">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="c3b6b-702">閉じる()</span><span class="sxs-lookup"><span data-stu-id="c3b6b-702">close()</span></span>

<span data-ttu-id="c3b6b-703">作成中の現在のアイテムを閉じます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="c3b6b-p137">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="c3b6b-706">Outlook on the webでは、アイテムが予定で、`saveAsync`を用いて事前に保存されている場合、アイテムが最後に保存されてから何も変更されていない場合でも、ユーザーに対して保存、破棄または取り消すよう求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="c3b6b-707">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-708">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-708">Requirements</span></span>

|<span data-ttu-id="c3b6b-709">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-709">Requirement</span></span>| <span data-ttu-id="c3b6b-710">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-711">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-711">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-712">1.3</span><span class="sxs-lookup"><span data-stu-id="c3b6b-712">1.3</span></span>|
|[<span data-ttu-id="c3b6b-713">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-713">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-714">制限あり</span><span class="sxs-lookup"><span data-stu-id="c3b6b-714">Restricted</span></span>|
|[<span data-ttu-id="c3b6b-715">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-715">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-716">作成</span><span class="sxs-lookup"><span data-stu-id="c3b6b-716">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="c3b6b-717">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-717">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="c3b6b-718">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c3b6b-719">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-719">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c3b6b-720">Outlook Web App では、回答フォームは、3 列ビューのポップアウト形式、および 2 列または 1 列ビューのポップアップ フォームで表示されます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-720">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c3b6b-721">文字列パラメータのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="c3b6b-p138">`formData.attachments` パラメータで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3b6b-725">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="c3b6b-725">Parameters:</span></span>

| <span data-ttu-id="c3b6b-726">名前</span><span class="sxs-lookup"><span data-stu-id="c3b6b-726">Name</span></span> | <span data-ttu-id="c3b6b-727">型</span><span class="sxs-lookup"><span data-stu-id="c3b6b-727">Type</span></span> | <span data-ttu-id="c3b6b-728">属性</span><span class="sxs-lookup"><span data-stu-id="c3b6b-728">Attributes</span></span> | <span data-ttu-id="c3b6b-729">説明</span><span class="sxs-lookup"><span data-stu-id="c3b6b-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="c3b6b-730">文字列 |オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3b6b-730">String &#124; Object</span></span>| |<span data-ttu-id="c3b6b-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c3b6b-733">**または**</span><span class="sxs-lookup"><span data-stu-id="c3b6b-733">**OR**</span></span><br/><span data-ttu-id="c3b6b-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクトです。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="c3b6b-736">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-736">String</span></span> | <span data-ttu-id="c3b6b-737">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-737">&lt;optional&gt;</span></span> | <span data-ttu-id="c3b6b-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="c3b6b-740">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="c3b6b-741">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-741">&lt;optional&gt;</span></span> | <span data-ttu-id="c3b6b-742">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="c3b6b-743">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-743">String</span></span> | | <span data-ttu-id="c3b6b-p142">添付ファイルの種類を示します。添付ファイルの場合は `file`、添付アイテムの場合は `item` でなければなりません。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="c3b6b-746">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-746">String</span></span> | | <span data-ttu-id="c3b6b-747">添付ファイル名を含む文字列で、最長 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="c3b6b-748">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-748">String</span></span> | | <span data-ttu-id="c3b6b-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URIです。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="c3b6b-751">ブール値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-751">Boolean</span></span> | | <span data-ttu-id="c3b6b-p144">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="c3b6b-754">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-754">String</span></span> | | <span data-ttu-id="c3b6b-p145">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムIDです。最大長が 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="c3b6b-758">関数</span><span class="sxs-lookup"><span data-stu-id="c3b6b-758">function</span></span> | <span data-ttu-id="c3b6b-759">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-759">&lt;optional&gt;</span></span> | <span data-ttu-id="c3b6b-760">メソッドが完了すると、 `callback` パラメータに渡された関数が、シングル パラメータ, `asyncResult`で呼び出されます。これは [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c3b6b-761">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-761">Requirements</span></span>

|<span data-ttu-id="c3b6b-762">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-762">Requirement</span></span>| <span data-ttu-id="c3b6b-763">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-764">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-764">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-765">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-765">1.0</span></span>|
|[<span data-ttu-id="c3b6b-766">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-766">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-767">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-768">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-768">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-769">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c3b6b-770">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-770">Examples</span></span>

<span data-ttu-id="c3b6b-771">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="c3b6b-772">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-772">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="c3b6b-773">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-773">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c3b6b-774">本文と添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-774">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c3b6b-775">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-775">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c3b6b-776">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="c3b6b-777">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-777">displayReplyForm(formData)</span></span>

<span data-ttu-id="c3b6b-778">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c3b6b-779">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-779">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c3b6b-780">Outlook Web App では、回答フォームは、3 列ビューのポップアウト形式、および 2 列または 1 列ビューのポップアップ フォームで表示されます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-780">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c3b6b-781">文字列パラメータのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="c3b6b-p146">`formData.attachments` パラメータで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3b6b-785">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="c3b6b-785">Parameters:</span></span>

| <span data-ttu-id="c3b6b-786">名前</span><span class="sxs-lookup"><span data-stu-id="c3b6b-786">Name</span></span> | <span data-ttu-id="c3b6b-787">型</span><span class="sxs-lookup"><span data-stu-id="c3b6b-787">Type</span></span> | <span data-ttu-id="c3b6b-788">属性</span><span class="sxs-lookup"><span data-stu-id="c3b6b-788">Attributes</span></span> | <span data-ttu-id="c3b6b-789">説明</span><span class="sxs-lookup"><span data-stu-id="c3b6b-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="c3b6b-790">文字列 |オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3b6b-790">String &#124; Object</span></span>| | <span data-ttu-id="c3b6b-p147">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c3b6b-793">**または**</span><span class="sxs-lookup"><span data-stu-id="c3b6b-793">**OR**</span></span><br/><span data-ttu-id="c3b6b-p148">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクトです。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="c3b6b-796">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-796">String</span></span> | <span data-ttu-id="c3b6b-797">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-797">&lt;optional&gt;</span></span> | <span data-ttu-id="c3b6b-p149">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="c3b6b-800">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="c3b6b-801">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-801">&lt;optional&gt;</span></span> | <span data-ttu-id="c3b6b-802">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="c3b6b-803">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-803">String</span></span> | | <span data-ttu-id="c3b6b-p150">添付ファイルの種類を示します。添付ファイルの場合は `file`、添付アイテムの場合は `item` でなければなりません。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="c3b6b-806">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-806">String</span></span> | | <span data-ttu-id="c3b6b-807">添付ファイル名を含む文字列で、最長 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="c3b6b-808">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-808">String</span></span> | | <span data-ttu-id="c3b6b-p151">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URIです。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="c3b6b-811">ブール値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-811">Boolean</span></span> | | <span data-ttu-id="c3b6b-p152">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="c3b6b-814">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-814">String</span></span> | | <span data-ttu-id="c3b6b-p153">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムIDです。最大長が 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="c3b6b-818">関数</span><span class="sxs-lookup"><span data-stu-id="c3b6b-818">function</span></span> | <span data-ttu-id="c3b6b-819">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-819">&lt;optional&gt;</span></span> | <span data-ttu-id="c3b6b-820">メソッドが完了すると、 `callback` パラメータに渡された関数が、シングル パラメータ, `asyncResult`で呼び出されます。これは [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c3b6b-821">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-821">Requirements</span></span>

|<span data-ttu-id="c3b6b-822">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-822">Requirement</span></span>| <span data-ttu-id="c3b6b-823">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-824">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-824">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-825">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-825">1.0</span></span>|
|[<span data-ttu-id="c3b6b-826">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-826">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-827">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-828">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-828">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-829">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c3b6b-830">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-830">Examples</span></span>

<span data-ttu-id="c3b6b-831">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="c3b6b-832">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-832">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="c3b6b-833">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-833">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c3b6b-834">本文と添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-834">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c3b6b-835">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-835">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c3b6b-836">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="c3b6b-837">getEntities() → {[エンティティ](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="c3b6b-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="c3b6b-838">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-838">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="c3b6b-839">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-839">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-840">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-840">Requirements</span></span>

|<span data-ttu-id="c3b6b-841">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-841">Requirement</span></span>| <span data-ttu-id="c3b6b-842">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-843">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-843">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-844">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-844">1.0</span></span>|
|[<span data-ttu-id="c3b6b-845">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-845">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-846">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-847">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-847">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-848">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3b6b-849">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="c3b6b-849">Returns:</span></span>

<span data-ttu-id="c3b6b-850">型:[エンティティ](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-850">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="c3b6b-851">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-851">Example</span></span>

<span data-ttu-id="c3b6b-852">次の例では、現在のアイテムの本文内の連絡先のエンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-852">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="c3b6b-853">getEntitiesByType(entityType)] → [(Null 許容) {<(String|[連絡先](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[電話番号](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion)) >}</span><span class="sxs-lookup"><span data-stu-id="c3b6b-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c3b6b-854">選択したアイテム内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-854">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="c3b6b-855">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-855">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3b6b-856">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="c3b6b-856">Parameters:</span></span>

|<span data-ttu-id="c3b6b-857">名前</span><span class="sxs-lookup"><span data-stu-id="c3b6b-857">Name</span></span>| <span data-ttu-id="c3b6b-858">型</span><span class="sxs-lookup"><span data-stu-id="c3b6b-858">Type</span></span>| <span data-ttu-id="c3b6b-859">説明</span><span class="sxs-lookup"><span data-stu-id="c3b6b-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="c3b6b-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="c3b6b-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="c3b6b-861">EntityType 列挙値の 1 つです。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3b6b-862">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-862">Requirements</span></span>

|<span data-ttu-id="c3b6b-863">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-863">Requirement</span></span>| <span data-ttu-id="c3b6b-864">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-865">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-865">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-866">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-866">1.0</span></span>|
|[<span data-ttu-id="c3b6b-867">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-867">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-868">制限あり</span><span class="sxs-lookup"><span data-stu-id="c3b6b-868">Restricted</span></span>|
|[<span data-ttu-id="c3b6b-869">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-869">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-870">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3b6b-871">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="c3b6b-871">Returns:</span></span>

<span data-ttu-id="c3b6b-872">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは 空白を返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="c3b6b-873">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-873">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="c3b6b-874">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメータ内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="c3b6b-875">このメソッドを使用する最小限のアクセス許可レベルは **制限あり** ですが、一部のエンティティ型には、次のテーブルで指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="c3b6b-876">の値 `entityType`</span><span class="sxs-lookup"><span data-stu-id="c3b6b-876">Value of `entityType`</span></span> | <span data-ttu-id="c3b6b-877">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="c3b6b-877">Type of objects in returned array</span></span> | <span data-ttu-id="c3b6b-878">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="c3b6b-879">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-879">String</span></span> | <span data-ttu-id="c3b6b-880">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="c3b6b-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="c3b6b-881">連絡先</span><span class="sxs-lookup"><span data-stu-id="c3b6b-881">Contact</span></span> | <span data-ttu-id="c3b6b-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c3b6b-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="c3b6b-883">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-883">String</span></span> | <span data-ttu-id="c3b6b-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c3b6b-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="c3b6b-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="c3b6b-885">MeetingSuggestion</span></span> | <span data-ttu-id="c3b6b-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c3b6b-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="c3b6b-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="c3b6b-887">PhoneNumber</span></span> | <span data-ttu-id="c3b6b-888">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="c3b6b-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="c3b6b-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="c3b6b-889">TaskSuggestion</span></span> | <span data-ttu-id="c3b6b-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c3b6b-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="c3b6b-891">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-891">String</span></span> | <span data-ttu-id="c3b6b-892">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="c3b6b-892">**Restricted**</span></span> |

<span data-ttu-id="c3b6b-893">型:Array.<(文字列|[連絡先](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c3b6b-893">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="c3b6b-894">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-894">Example</span></span>

<span data-ttu-id="c3b6b-895">次の例は、現在のアイテムの本文にある郵便アドレスを表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-895">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="c3b6b-896">getFilteredEntitiesByName(name)] → [(Null 許容空白が可能) {<(文字列|[連絡先](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[電話番号](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion)) >}</span><span class="sxs-lookup"><span data-stu-id="c3b6b-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c3b6b-897">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c3b6b-898">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-898">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c3b6b-899">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3b6b-900">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="c3b6b-900">Parameters:</span></span>

|<span data-ttu-id="c3b6b-901">名前</span><span class="sxs-lookup"><span data-stu-id="c3b6b-901">Name</span></span>| <span data-ttu-id="c3b6b-902">型</span><span class="sxs-lookup"><span data-stu-id="c3b6b-902">Type</span></span>| <span data-ttu-id="c3b6b-903">説明</span><span class="sxs-lookup"><span data-stu-id="c3b6b-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="c3b6b-904">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-904">String</span></span>|<span data-ttu-id="c3b6b-905">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3b6b-906">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-906">Requirements</span></span>

|<span data-ttu-id="c3b6b-907">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-907">Requirement</span></span>| <span data-ttu-id="c3b6b-908">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-909">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-909">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-910">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-910">1.0</span></span>|
|[<span data-ttu-id="c3b6b-911">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-911">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-912">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-913">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-913">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-914">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3b6b-915">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="c3b6b-915">Returns:</span></span>

<span data-ttu-id="c3b6b-p155">`ItemHasKnownEntity` 要素が `name` パラメータと一致する`FilterName` 要素値を持つマニフェスト内に ない場合、メソッドは `null`を返します。  `name` パラメータがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="c3b6b-918">型:Array.<(文字列|[連絡先](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c3b6b-918">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="c3b6b-919">getRegExMatches() → {オブジェクト}</span><span class="sxs-lookup"><span data-stu-id="c3b6b-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="c3b6b-920">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c3b6b-921">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-921">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c3b6b-p156">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c3b6b-925">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c3b6b-926">`getRegExMatches` から返されるオブジェクトには、 `fruits` および `veggies` という 2 つのプロパティがありえます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c3b6b-p157">アイテムの 本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。代わりに [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-930">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-930">Requirements</span></span>

|<span data-ttu-id="c3b6b-931">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-931">Requirement</span></span>| <span data-ttu-id="c3b6b-932">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-933">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-933">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-934">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-934">1.0</span></span>|
|[<span data-ttu-id="c3b6b-935">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-935">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-936">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-937">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-937">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-938">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3b6b-939">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="c3b6b-939">Returns:</span></span>

<span data-ttu-id="c3b6b-p158">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクトです。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="c3b6b-942">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c3b6b-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c3b6b-943">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3b6b-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c3b6b-944">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-944">Example</span></span>

<span data-ttu-id="c3b6b-945">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="c3b6b-946">getRegExMatchesByName(name)] → [(Null 許容) {配列.< 文字列 >}</span><span class="sxs-lookup"><span data-stu-id="c3b6b-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="c3b6b-947">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c3b6b-948">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-948">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c3b6b-949">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="c3b6b-p159">アイテムの 本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3b6b-952">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="c3b6b-952">Parameters:</span></span>

|<span data-ttu-id="c3b6b-953">名前</span><span class="sxs-lookup"><span data-stu-id="c3b6b-953">Name</span></span>| <span data-ttu-id="c3b6b-954">型</span><span class="sxs-lookup"><span data-stu-id="c3b6b-954">Type</span></span>| <span data-ttu-id="c3b6b-955">説明</span><span class="sxs-lookup"><span data-stu-id="c3b6b-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="c3b6b-956">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-956">String</span></span>|<span data-ttu-id="c3b6b-957">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3b6b-958">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-958">Requirements</span></span>

|<span data-ttu-id="c3b6b-959">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-959">Requirement</span></span>| <span data-ttu-id="c3b6b-960">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-961">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-961">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-962">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-962">1.0</span></span>|
|[<span data-ttu-id="c3b6b-963">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-963">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-964">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-965">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-965">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-966">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3b6b-967">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="c3b6b-967">Returns:</span></span>

<span data-ttu-id="c3b6b-968">マニフェスト XML ファイルで定義された正規表現に一致する文字列を含む配列です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="c3b6b-969">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c3b6b-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c3b6b-970">配列. < 文字列 ></span><span class="sxs-lookup"><span data-stu-id="c3b6b-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c3b6b-971">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-971">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="c3b6b-972">getSelectedDataAsync (coercionType、[オプション] 、コールバック) →{文字列}</span><span class="sxs-lookup"><span data-stu-id="c3b6b-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="c3b6b-973">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="c3b6b-p160">選択したデータがなく、カーソルが本文または件名にある場合、メソッドは選択したデータに対して null が返します。本文または件名以外のフィールドが選択されている場合、メソッドは`InvalidSelection` エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3b6b-976">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="c3b6b-976">Parameters:</span></span>

|<span data-ttu-id="c3b6b-977">名前</span><span class="sxs-lookup"><span data-stu-id="c3b6b-977">Name</span></span>| <span data-ttu-id="c3b6b-978">型</span><span class="sxs-lookup"><span data-stu-id="c3b6b-978">Type</span></span>| <span data-ttu-id="c3b6b-979">属性</span><span class="sxs-lookup"><span data-stu-id="c3b6b-979">Attributes</span></span>| <span data-ttu-id="c3b6b-980">説明</span><span class="sxs-lookup"><span data-stu-id="c3b6b-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="c3b6b-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c3b6b-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="c3b6b-p161">データの形式を要求します。携帯ショートメール（SMS）の場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="c3b6b-985">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3b6b-985">Object</span></span>| <span data-ttu-id="c3b6b-986">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-986">&lt;optional&gt;</span></span>|<span data-ttu-id="c3b6b-987">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c3b6b-988">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3b6b-988">Object</span></span>| <span data-ttu-id="c3b6b-989">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-989">&lt;optional&gt;</span></span>|<span data-ttu-id="c3b6b-990">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c3b6b-991">関数</span><span class="sxs-lookup"><span data-stu-id="c3b6b-991">function</span></span>||<span data-ttu-id="c3b6b-992">メソッドが完了すると、`callback`パラメータに渡された関数が、シングル パラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c3b6b-993">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="c3b6b-994">選択元のソース プロパティにアクセスするには、 `asyncResult.value.sourceProperty` を呼び出します。これは `body`   または `subject`   になります。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-994">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3b6b-995">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-995">Requirements</span></span>

|<span data-ttu-id="c3b6b-996">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-996">Requirement</span></span>| <span data-ttu-id="c3b6b-997">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-998">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-998">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-999">1.2</span><span class="sxs-lookup"><span data-stu-id="c3b6b-999">1.2</span></span>|
|[<span data-ttu-id="c3b6b-1000">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1000">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3b6b-1002">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1002">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-1003">作成</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3b6b-1004">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1004">Returns:</span></span>

<span data-ttu-id="c3b6b-1005">`coercionType` で決定された書式設定の文字列としての選択されたデータ</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="c3b6b-1006">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c3b6b-1007">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c3b6b-1008">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1008">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="c3b6b-1009">getSelectedEntities() → {[エンティティ](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="c3b6b-p163">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p163">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="c3b6b-1012">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1012">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-1013">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1013">Requirements</span></span>

|<span data-ttu-id="c3b6b-1014">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1014">Requirement</span></span>| <span data-ttu-id="c3b6b-1015">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-1016">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1016">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1017">-16</span></span> |
|[<span data-ttu-id="c3b6b-1018">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1018">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1019">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-1020">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1020">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-1021">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3b6b-1022">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1022">Returns:</span></span>

<span data-ttu-id="c3b6b-1023">型:[エンティティ](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1023">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="c3b6b-1024">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1024">Example</span></span>

<span data-ttu-id="c3b6b-1025">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="c3b6b-1026">getSelectedRegExMatches() → {オブジェクト}</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="c3b6b-p164">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="c3b6b-1029">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1029">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c3b6b-p165">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c3b6b-1033">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c3b6b-1034">`getRegExMatches` から返されるオブジェクトには、 `fruits` および `veggies` という 2 つのプロパティがありえます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c3b6b-p166">アイテムの 本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。代わりに [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3b6b-1038">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1038">Requirements</span></span>

|<span data-ttu-id="c3b6b-1039">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1039">Requirement</span></span>| <span data-ttu-id="c3b6b-1040">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-1041">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1041">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1042">-16</span></span> |
|[<span data-ttu-id="c3b6b-1043">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1043">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1044">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-1045">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1045">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-1046">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3b6b-1047">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1047">Returns:</span></span>

<span data-ttu-id="c3b6b-p167">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクトです。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="c3b6b-1050">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1050">Example</span></span>

<span data-ttu-id="c3b6b-1051">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="c3b6b-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="c3b6b-1053">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="c3b6b-p168">カスタム プロパティは、アプリケーションごと、アイテムごとにキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3b6b-1057">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1057">Parameters:</span></span>

|<span data-ttu-id="c3b6b-1058">名前</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1058">Name</span></span>| <span data-ttu-id="c3b6b-1059">型</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1059">Type</span></span>| <span data-ttu-id="c3b6b-1060">属性</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1060">Attributes</span></span>| <span data-ttu-id="c3b6b-1061">説明</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="c3b6b-1062">関数</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1062">function</span></span>||<span data-ttu-id="c3b6b-1063">メソッドが完了すると、`callback`パラメータに渡された関数が、シングル パラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c3b6b-1064">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="c3b6b-1065">アイテムからカスタム プロパティを取得、設定、削除して、サーバーへのカスタム プロパティ セット バックに対する変更を保存するのに、このオブジェクトが使用できます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1065">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="c3b6b-1066">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1066">Object</span></span>| <span data-ttu-id="c3b6b-1067">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="c3b6b-1068">開発者は、コールバック 関数でアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1068">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="c3b6b-1069">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3b6b-1070">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1070">Requirements</span></span>

|<span data-ttu-id="c3b6b-1071">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1071">Requirement</span></span>| <span data-ttu-id="c3b6b-1072">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-1073">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1073">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-1074">1.0</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1074">1.0</span></span>|
|[<span data-ttu-id="c3b6b-1075">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1075">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1076">ReadItem</span></span>|
|[<span data-ttu-id="c3b6b-1077">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1077">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-1078">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1078">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-1079">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1079">Example</span></span>

<span data-ttu-id="c3b6b-p171">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="c3b6b-1083">removeAttachmentAsync (attachmentId、[オプション]、[コールバック])</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="c3b6b-1084">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="c3b6b-p172">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3b6b-1089">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1089">Parameters:</span></span>

|<span data-ttu-id="c3b6b-1090">名前</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1090">Name</span></span>| <span data-ttu-id="c3b6b-1091">型</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1091">Type</span></span>| <span data-ttu-id="c3b6b-1092">属性</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1092">Attributes</span></span>| <span data-ttu-id="c3b6b-1093">説明</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="c3b6b-1094">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1094">String</span></span>||<span data-ttu-id="c3b6b-p173">削除する添付ファイルの識別子です。文字列の最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p173">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="c3b6b-1097">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1097">Object</span></span>| <span data-ttu-id="c3b6b-1098">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="c3b6b-1099">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1099">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c3b6b-1100">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1100">Object</span></span>| <span data-ttu-id="c3b6b-1101">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1101">&lt;optional&gt;</span></span>|<span data-ttu-id="c3b6b-1102">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1102">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c3b6b-1103">関数</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1103">function</span></span>| <span data-ttu-id="c3b6b-1104">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1104">&lt;optional&gt;</span></span>|<span data-ttu-id="c3b6b-1105">メソッドが完了すると、`callback`パラメータに渡された関数が、シングル パラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1105">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c3b6b-1106">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1106">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c3b6b-1107">エラー</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1107">Errors</span></span>

| <span data-ttu-id="c3b6b-1108">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1108">Error code</span></span> | <span data-ttu-id="c3b6b-1109">説明</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1109">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="c3b6b-1110">添付ファイルの識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1110">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c3b6b-1111">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1111">Requirements</span></span>

|<span data-ttu-id="c3b6b-1112">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1112">Requirement</span></span>| <span data-ttu-id="c3b6b-1113">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1113">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-1114">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1114">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-1115">1.1</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1115">1.1</span></span>|
|[<span data-ttu-id="c3b6b-1116">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1116">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-1117">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1117">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3b6b-1118">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1118">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-1119">作成</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1119">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-1120">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1120">Example</span></span>

<span data-ttu-id="c3b6b-1121">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1121">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="c3b6b-1122">saveAsync ([オプション] 、コールバック)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1122">saveAsync([options], callback)</span></span>

<span data-ttu-id="c3b6b-1123">アイテムを非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1123">Asynchronously saves an item.</span></span>

<span data-ttu-id="c3b6b-p174">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッド経由でアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p174">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="c3b6b-1127">アドインが、WS または REST API を使用しようとして `itemId` を取得するために、新規作成モードでアイテム上の `saveAsync` を呼び出す場合、Outlook キャッシュ モードでは、アイテムがサーバーと実際に同期するまでに時間がかかる場合があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1127">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="c3b6b-1128">アイテムが同期されるまで、 `itemId` を使用すると、エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1128">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="c3b6b-p176">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p176">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="c3b6b-1132">次のクライアントは、作成モードで予定上の `saveAsync` に対して様々なふるまいをします。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1132">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="c3b6b-1133">Mac Outlook は、新規作成モードの会議場で `saveAsync` をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1133">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="c3b6b-1134">Mac Outlookの会議場で  `saveAsync` を呼びだすと、エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1134">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="c3b6b-1135">新規作成モードで予定上で `saveAsync` が呼び出されると、Outlook on the webは常に、招待状または更新を送信します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1135">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3b6b-1136">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1136">Parameters:</span></span>

|<span data-ttu-id="c3b6b-1137">名前</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1137">Name</span></span>| <span data-ttu-id="c3b6b-1138">型</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1138">Type</span></span>| <span data-ttu-id="c3b6b-1139">属性</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1139">Attributes</span></span>| <span data-ttu-id="c3b6b-1140">説明</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1140">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="c3b6b-1141">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1141">Object</span></span>| <span data-ttu-id="c3b6b-1142">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="c3b6b-1143">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1143">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c3b6b-1144">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1144">Object</span></span>| <span data-ttu-id="c3b6b-1145">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1145">&lt;optional&gt;</span></span>|<span data-ttu-id="c3b6b-1146">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1146">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c3b6b-1147">関数</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1147">function</span></span>||<span data-ttu-id="c3b6b-1148">メソッドが完了すると、`callback`パラメータに渡された関数が、シングル パラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1148">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c3b6b-1149">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1149">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3b6b-1150">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1150">Requirements</span></span>

|<span data-ttu-id="c3b6b-1151">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1151">Requirement</span></span>| <span data-ttu-id="c3b6b-1152">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1152">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-1153">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1153">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-1154">1.3</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1154">1.3</span></span>|
|[<span data-ttu-id="c3b6b-1155">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-1156">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1156">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3b6b-1157">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-1158">作成</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1158">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c3b6b-1159">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1159">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="c3b6b-p178">次の例は、コールバック関数に渡される `result` パラメータの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p178">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="c3b6b-1162">setSelectedDataAsync (データ、[オプション] 、コールバック)</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1162">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="c3b6b-1163">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1163">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="c3b6b-p179">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p179">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3b6b-1167">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1167">Parameters:</span></span>

|<span data-ttu-id="c3b6b-1168">名前</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1168">Name</span></span>| <span data-ttu-id="c3b6b-1169">型</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1169">Type</span></span>| <span data-ttu-id="c3b6b-1170">属性</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1170">Attributes</span></span>| <span data-ttu-id="c3b6b-1171">説明</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1171">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="c3b6b-1172">文字列</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1172">String</span></span>||<span data-ttu-id="c3b6b-p180">挿入されるデータです。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p180">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="c3b6b-1176">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1176">Object</span></span>| <span data-ttu-id="c3b6b-1177">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="c3b6b-1178">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1178">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c3b6b-1179">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1179">Object</span></span>| <span data-ttu-id="c3b6b-1180">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="c3b6b-1181">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1181">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="c3b6b-1182">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1182">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="c3b6b-1183">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1183">&lt;optional&gt;</span></span>|<span data-ttu-id="c3b6b-p181">`text` の場合、Outlook Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p181">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="c3b6b-p182">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-p182">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="c3b6b-1188">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1188">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="c3b6b-1189">関数</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1189">function</span></span>||<span data-ttu-id="c3b6b-1190">メソッドが完了すると、`callback`パラメータに渡された関数が、シングル パラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1190">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c3b6b-1191">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1191">Requirements</span></span>

|<span data-ttu-id="c3b6b-1192">要件</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1192">Requirement</span></span>| <span data-ttu-id="c3b6b-1193">値</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3b6b-1194">メールボックスの最低要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1194">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3b6b-1195">1.2</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1195">1.2</span></span>|
|[<span data-ttu-id="c3b6b-1196">最低限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3b6b-1197">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1197">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3b6b-1198">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3b6b-1199">作成</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1199">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c3b6b-1200">例</span><span class="sxs-lookup"><span data-stu-id="c3b6b-1200">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```