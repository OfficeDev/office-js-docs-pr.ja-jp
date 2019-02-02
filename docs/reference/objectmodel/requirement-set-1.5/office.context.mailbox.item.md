---
title: Office.context.mailbox.item - requirement set 1.5
description: ''
ms.date: 01/30/2019
localization_priority: Priority
ms.openlocfilehash: cca0bb4baa15d72a58909ca1417eb52a9bf70a8f
ms.sourcegitcommit: bf5c56d9b8c573e42bf2268e10ca3fd4d2bb4ff9
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/01/2019
ms.locfileid: "29701842"
---
# <a name="item"></a><span data-ttu-id="88e57-102">item</span><span class="sxs-lookup"><span data-stu-id="88e57-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="88e57-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="88e57-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="88e57-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="88e57-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="88e57-106">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-106">Requirements</span></span>

|<span data-ttu-id="88e57-107">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-107">Requirement</span></span>| <span data-ttu-id="88e57-108">値</span><span class="sxs-lookup"><span data-stu-id="88e57-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-110">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-110">1.0</span></span>|
|[<span data-ttu-id="88e57-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="88e57-112">Restricted</span></span>|
|[<span data-ttu-id="88e57-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-114">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="88e57-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="88e57-115">Members and methods</span></span>

| <span data-ttu-id="88e57-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-116">Member</span></span> | <span data-ttu-id="88e57-117">種類</span><span class="sxs-lookup"><span data-stu-id="88e57-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="88e57-118">attachments</span><span class="sxs-lookup"><span data-stu-id="88e57-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails) | <span data-ttu-id="88e57-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-119">Member</span></span> |
| [<span data-ttu-id="88e57-120">bcc</span><span class="sxs-lookup"><span data-stu-id="88e57-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="88e57-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-121">Member</span></span> |
| [<span data-ttu-id="88e57-122">body</span><span class="sxs-lookup"><span data-stu-id="88e57-122">body</span></span>](#body-bodyjavascriptapioutlook15officebody) | <span data-ttu-id="88e57-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-123">Member</span></span> |
| [<span data-ttu-id="88e57-124">cc</span><span class="sxs-lookup"><span data-stu-id="88e57-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="88e57-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-125">Member</span></span> |
| [<span data-ttu-id="88e57-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="88e57-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="88e57-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-127">Member</span></span> |
| [<span data-ttu-id="88e57-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="88e57-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="88e57-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-129">Member</span></span> |
| [<span data-ttu-id="88e57-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="88e57-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="88e57-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-131">Member</span></span> |
| [<span data-ttu-id="88e57-132">end</span><span class="sxs-lookup"><span data-stu-id="88e57-132">end</span></span>](#end-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="88e57-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-133">Member</span></span> |
| [<span data-ttu-id="88e57-134">from</span><span class="sxs-lookup"><span data-stu-id="88e57-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="88e57-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-135">Member</span></span> |
| [<span data-ttu-id="88e57-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="88e57-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="88e57-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-137">Member</span></span> |
| [<span data-ttu-id="88e57-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="88e57-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="88e57-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-139">Member</span></span> |
| [<span data-ttu-id="88e57-140">itemId</span><span class="sxs-lookup"><span data-stu-id="88e57-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="88e57-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-141">Member</span></span> |
| [<span data-ttu-id="88e57-142">itemType</span><span class="sxs-lookup"><span data-stu-id="88e57-142">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) | <span data-ttu-id="88e57-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-143">Member</span></span> |
| [<span data-ttu-id="88e57-144">location</span><span class="sxs-lookup"><span data-stu-id="88e57-144">location</span></span>](#location-stringlocationjavascriptapioutlook15officelocation) | <span data-ttu-id="88e57-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-145">Member</span></span> |
| [<span data-ttu-id="88e57-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="88e57-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="88e57-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-147">Member</span></span> |
| [<span data-ttu-id="88e57-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="88e57-148">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages) | <span data-ttu-id="88e57-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-149">Member</span></span> |
| [<span data-ttu-id="88e57-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="88e57-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="88e57-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-151">Member</span></span> |
| [<span data-ttu-id="88e57-152">organizer</span><span class="sxs-lookup"><span data-stu-id="88e57-152">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="88e57-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-153">Member</span></span> |
| [<span data-ttu-id="88e57-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="88e57-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="88e57-155">Member</span><span class="sxs-lookup"><span data-stu-id="88e57-155">Member</span></span> |
| [<span data-ttu-id="88e57-156">sender</span><span class="sxs-lookup"><span data-stu-id="88e57-156">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="88e57-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-157">Member</span></span> |
| [<span data-ttu-id="88e57-158">start</span><span class="sxs-lookup"><span data-stu-id="88e57-158">start</span></span>](#start-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="88e57-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-159">Member</span></span> |
| [<span data-ttu-id="88e57-160">subject</span><span class="sxs-lookup"><span data-stu-id="88e57-160">subject</span></span>](#subject-stringsubjectjavascriptapioutlook15officesubject) | <span data-ttu-id="88e57-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-161">Member</span></span> |
| [<span data-ttu-id="88e57-162">to</span><span class="sxs-lookup"><span data-stu-id="88e57-162">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="88e57-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-163">Member</span></span> |
| [<span data-ttu-id="88e57-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="88e57-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="88e57-165">メソッド</span><span class="sxs-lookup"><span data-stu-id="88e57-165">Method</span></span> |
| [<span data-ttu-id="88e57-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="88e57-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="88e57-167">メソッド</span><span class="sxs-lookup"><span data-stu-id="88e57-167">Method</span></span> |
| [<span data-ttu-id="88e57-168">close</span><span class="sxs-lookup"><span data-stu-id="88e57-168">close</span></span>](#close) | <span data-ttu-id="88e57-169">メソッド</span><span class="sxs-lookup"><span data-stu-id="88e57-169">Method</span></span> |
| [<span data-ttu-id="88e57-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="88e57-170">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="88e57-171">メソッド</span><span class="sxs-lookup"><span data-stu-id="88e57-171">Method</span></span> |
| [<span data-ttu-id="88e57-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="88e57-172">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="88e57-173">メソッド</span><span class="sxs-lookup"><span data-stu-id="88e57-173">Method</span></span> |
| [<span data-ttu-id="88e57-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="88e57-174">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook15officeentities) | <span data-ttu-id="88e57-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="88e57-175">Method</span></span> |
| [<span data-ttu-id="88e57-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="88e57-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="88e57-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="88e57-177">Method</span></span> |
| [<span data-ttu-id="88e57-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="88e57-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="88e57-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="88e57-179">Method</span></span> |
| [<span data-ttu-id="88e57-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="88e57-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="88e57-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="88e57-181">Method</span></span> |
| [<span data-ttu-id="88e57-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="88e57-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="88e57-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="88e57-183">Method</span></span> |
| [<span data-ttu-id="88e57-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="88e57-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="88e57-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="88e57-185">Method</span></span> |
| [<span data-ttu-id="88e57-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="88e57-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="88e57-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="88e57-187">Method</span></span> |
| [<span data-ttu-id="88e57-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="88e57-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="88e57-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="88e57-189">Method</span></span> |
| [<span data-ttu-id="88e57-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="88e57-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="88e57-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="88e57-191">Method</span></span> |
| [<span data-ttu-id="88e57-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="88e57-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="88e57-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="88e57-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="88e57-194">例</span><span class="sxs-lookup"><span data-stu-id="88e57-194">Example</span></span>

<span data-ttu-id="88e57-195">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="88e57-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="88e57-196">メンバー</span><span class="sxs-lookup"><span data-stu-id="88e57-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="88e57-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="88e57-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="88e57-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="88e57-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="88e57-200">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="88e57-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="88e57-201">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="88e57-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-202">種類:</span><span class="sxs-lookup"><span data-stu-id="88e57-202">Type:</span></span>

*   <span data-ttu-id="88e57-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="88e57-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="88e57-204">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-204">Requirements</span></span>

|<span data-ttu-id="88e57-205">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-205">Requirement</span></span>| <span data-ttu-id="88e57-206">値</span><span class="sxs-lookup"><span data-stu-id="88e57-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-207">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-208">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-208">1.0</span></span>|
|[<span data-ttu-id="88e57-209">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-209">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-210">ReadItem</span></span>|
|[<span data-ttu-id="88e57-211">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-211">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-212">読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-213">例</span><span class="sxs-lookup"><span data-stu-id="88e57-213">Example</span></span>

<span data-ttu-id="88e57-214">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="88e57-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="88e57-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="88e57-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="88e57-216">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="88e57-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="88e57-217">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="88e57-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-218">型:</span><span class="sxs-lookup"><span data-stu-id="88e57-218">Type:</span></span>

*   [<span data-ttu-id="88e57-219">Recipients</span><span class="sxs-lookup"><span data-stu-id="88e57-219">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="88e57-220">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-220">Requirements</span></span>

|<span data-ttu-id="88e57-221">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-221">Requirement</span></span>| <span data-ttu-id="88e57-222">値</span><span class="sxs-lookup"><span data-stu-id="88e57-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-223">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-224">1.1</span><span class="sxs-lookup"><span data-stu-id="88e57-224">1.1</span></span>|
|[<span data-ttu-id="88e57-225">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-225">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-226">ReadItem</span></span>|
|[<span data-ttu-id="88e57-227">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-227">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-228">作成</span><span class="sxs-lookup"><span data-stu-id="88e57-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-229">例</span><span class="sxs-lookup"><span data-stu-id="88e57-229">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="88e57-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="88e57-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="88e57-231">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="88e57-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-232">型:</span><span class="sxs-lookup"><span data-stu-id="88e57-232">Type:</span></span>

*   [<span data-ttu-id="88e57-233">Body</span><span class="sxs-lookup"><span data-stu-id="88e57-233">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="88e57-234">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-234">Requirements</span></span>

|<span data-ttu-id="88e57-235">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-235">Requirement</span></span>| <span data-ttu-id="88e57-236">値</span><span class="sxs-lookup"><span data-stu-id="88e57-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-237">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-238">1.1</span><span class="sxs-lookup"><span data-stu-id="88e57-238">1.1</span></span>|
|[<span data-ttu-id="88e57-239">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-239">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-240">ReadItem</span></span>|
|[<span data-ttu-id="88e57-241">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-241">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-242">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-242">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="88e57-243">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="88e57-243">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="88e57-244">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="88e57-244">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="88e57-245">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="88e57-245">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="88e57-246">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="88e57-246">Read mode</span></span>

<span data-ttu-id="88e57-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="88e57-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="88e57-249">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="88e57-249">Compose mode</span></span>

<span data-ttu-id="88e57-250">`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-250">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-251">種類:</span><span class="sxs-lookup"><span data-stu-id="88e57-251">Type:</span></span>

*   <span data-ttu-id="88e57-252">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="88e57-252">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="88e57-253">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-253">Requirements</span></span>

|<span data-ttu-id="88e57-254">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-254">Requirement</span></span>| <span data-ttu-id="88e57-255">値</span><span class="sxs-lookup"><span data-stu-id="88e57-255">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-256">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-256">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-257">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-257">1.0</span></span>|
|[<span data-ttu-id="88e57-258">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-258">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-259">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-259">ReadItem</span></span>|
|[<span data-ttu-id="88e57-260">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-260">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-261">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-261">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-262">例</span><span class="sxs-lookup"><span data-stu-id="88e57-262">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="88e57-263">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="88e57-263">(nullable) conversationId :String</span></span>

<span data-ttu-id="88e57-264">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="88e57-264">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="88e57-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="88e57-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="88e57-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-269">型:</span><span class="sxs-lookup"><span data-stu-id="88e57-269">Type:</span></span>

*   <span data-ttu-id="88e57-270">String</span><span class="sxs-lookup"><span data-stu-id="88e57-270">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="88e57-271">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-271">Requirements</span></span>

|<span data-ttu-id="88e57-272">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-272">Requirement</span></span>| <span data-ttu-id="88e57-273">値</span><span class="sxs-lookup"><span data-stu-id="88e57-273">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-274">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-274">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-275">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-275">1.0</span></span>|
|[<span data-ttu-id="88e57-276">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-276">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-277">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-277">ReadItem</span></span>|
|[<span data-ttu-id="88e57-278">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-278">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-279">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-279">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="88e57-280">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="88e57-280">dateTimeCreated :Date</span></span>

<span data-ttu-id="88e57-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="88e57-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-283">型:</span><span class="sxs-lookup"><span data-stu-id="88e57-283">Type:</span></span>

*   <span data-ttu-id="88e57-284">日付</span><span class="sxs-lookup"><span data-stu-id="88e57-284">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="88e57-285">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-285">Requirements</span></span>

|<span data-ttu-id="88e57-286">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-286">Requirement</span></span>| <span data-ttu-id="88e57-287">値</span><span class="sxs-lookup"><span data-stu-id="88e57-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-288">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-289">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-289">1.0</span></span>|
|[<span data-ttu-id="88e57-290">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-291">ReadItem</span></span>|
|[<span data-ttu-id="88e57-292">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-293">読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-293">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-294">例</span><span class="sxs-lookup"><span data-stu-id="88e57-294">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="88e57-295">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="88e57-295">dateTimeModified :Date</span></span>

<span data-ttu-id="88e57-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="88e57-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="88e57-298">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="88e57-298">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-299">型:</span><span class="sxs-lookup"><span data-stu-id="88e57-299">Type:</span></span>

*   <span data-ttu-id="88e57-300">日付</span><span class="sxs-lookup"><span data-stu-id="88e57-300">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="88e57-301">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-301">Requirements</span></span>

|<span data-ttu-id="88e57-302">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-302">Requirement</span></span>| <span data-ttu-id="88e57-303">値</span><span class="sxs-lookup"><span data-stu-id="88e57-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-304">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-305">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-305">1.0</span></span>|
|[<span data-ttu-id="88e57-306">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-307">ReadItem</span></span>|
|[<span data-ttu-id="88e57-308">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-309">読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-310">例</span><span class="sxs-lookup"><span data-stu-id="88e57-310">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="88e57-311">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="88e57-311">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="88e57-312">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="88e57-312">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="88e57-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="88e57-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="88e57-315">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="88e57-315">Read mode</span></span>

<span data-ttu-id="88e57-316">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-316">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="88e57-317">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="88e57-317">Compose mode</span></span>

<span data-ttu-id="88e57-318">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-318">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="88e57-319">[`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="88e57-319">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-320">型:</span><span class="sxs-lookup"><span data-stu-id="88e57-320">Type:</span></span>

*   <span data-ttu-id="88e57-321">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="88e57-321">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="88e57-322">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-322">Requirements</span></span>

|<span data-ttu-id="88e57-323">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-323">Requirement</span></span>| <span data-ttu-id="88e57-324">値</span><span class="sxs-lookup"><span data-stu-id="88e57-324">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-325">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-325">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-326">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-326">1.0</span></span>|
|[<span data-ttu-id="88e57-327">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-327">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-328">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-328">ReadItem</span></span>|
|[<span data-ttu-id="88e57-329">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-329">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-330">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-330">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-331">例</span><span class="sxs-lookup"><span data-stu-id="88e57-331">Example</span></span>

<span data-ttu-id="88e57-332">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="88e57-332">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="88e57-333">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="88e57-333">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="88e57-p112">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="88e57-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="88e57-p113">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="88e57-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="88e57-338">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="88e57-338">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-339">種類:</span><span class="sxs-lookup"><span data-stu-id="88e57-339">Type:</span></span>

*   [<span data-ttu-id="88e57-340">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="88e57-340">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="88e57-341">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-341">Requirements</span></span>

|<span data-ttu-id="88e57-342">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-342">Requirement</span></span>| <span data-ttu-id="88e57-343">値</span><span class="sxs-lookup"><span data-stu-id="88e57-343">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-344">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-344">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-345">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-345">1.0</span></span>|
|[<span data-ttu-id="88e57-346">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-346">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-347">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-347">ReadItem</span></span>|
|[<span data-ttu-id="88e57-348">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-348">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-349">読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-349">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="88e57-350">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="88e57-350">internetMessageId :String</span></span>

<span data-ttu-id="88e57-p114">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="88e57-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-353">型:</span><span class="sxs-lookup"><span data-stu-id="88e57-353">Type:</span></span>

*   <span data-ttu-id="88e57-354">String</span><span class="sxs-lookup"><span data-stu-id="88e57-354">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="88e57-355">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-355">Requirements</span></span>

|<span data-ttu-id="88e57-356">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-356">Requirement</span></span>| <span data-ttu-id="88e57-357">値</span><span class="sxs-lookup"><span data-stu-id="88e57-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-358">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-358">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-359">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-359">1.0</span></span>|
|[<span data-ttu-id="88e57-360">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-361">ReadItem</span></span>|
|[<span data-ttu-id="88e57-362">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-363">読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-363">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-364">例</span><span class="sxs-lookup"><span data-stu-id="88e57-364">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="88e57-365">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="88e57-365">itemClass :String</span></span>

<span data-ttu-id="88e57-p115">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="88e57-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="88e57-p116">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="88e57-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="88e57-370">型</span><span class="sxs-lookup"><span data-stu-id="88e57-370">Type</span></span> | <span data-ttu-id="88e57-371">説明</span><span class="sxs-lookup"><span data-stu-id="88e57-371">Description</span></span> | <span data-ttu-id="88e57-372">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="88e57-372">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="88e57-373">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="88e57-373">Appointment items</span></span> | <span data-ttu-id="88e57-374">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="88e57-374">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="88e57-375">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="88e57-375">Message items</span></span> | <span data-ttu-id="88e57-376">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="88e57-376">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="88e57-377">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="88e57-377">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-378">型:</span><span class="sxs-lookup"><span data-stu-id="88e57-378">Type:</span></span>

*   <span data-ttu-id="88e57-379">String</span><span class="sxs-lookup"><span data-stu-id="88e57-379">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="88e57-380">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-380">Requirements</span></span>

|<span data-ttu-id="88e57-381">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-381">Requirement</span></span>| <span data-ttu-id="88e57-382">値</span><span class="sxs-lookup"><span data-stu-id="88e57-382">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-383">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-383">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-384">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-384">1.0</span></span>|
|[<span data-ttu-id="88e57-385">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-385">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-386">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-386">ReadItem</span></span>|
|[<span data-ttu-id="88e57-387">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-387">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-388">読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-388">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-389">例</span><span class="sxs-lookup"><span data-stu-id="88e57-389">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="88e57-390">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="88e57-390">(nullable) itemId :String</span></span>

<span data-ttu-id="88e57-p117">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="88e57-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="88e57-393">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="88e57-393">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="88e57-394">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="88e57-394">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="88e57-395">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="88e57-395">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="88e57-396">詳細は、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="88e57-396">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="88e57-p119">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-399">種類:</span><span class="sxs-lookup"><span data-stu-id="88e57-399">Type:</span></span>

*   <span data-ttu-id="88e57-400">String</span><span class="sxs-lookup"><span data-stu-id="88e57-400">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="88e57-401">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-401">Requirements</span></span>

|<span data-ttu-id="88e57-402">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-402">Requirement</span></span>| <span data-ttu-id="88e57-403">値</span><span class="sxs-lookup"><span data-stu-id="88e57-403">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-404">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-404">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-405">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-405">1.0</span></span>|
|[<span data-ttu-id="88e57-406">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-406">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-407">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-407">ReadItem</span></span>|
|[<span data-ttu-id="88e57-408">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-408">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-409">読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-409">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-410">例</span><span class="sxs-lookup"><span data-stu-id="88e57-410">Example</span></span>

<span data-ttu-id="88e57-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="88e57-413">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="88e57-413">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="88e57-414">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="88e57-414">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="88e57-415">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="88e57-415">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-416">型:</span><span class="sxs-lookup"><span data-stu-id="88e57-416">Type:</span></span>

*   [<span data-ttu-id="88e57-417">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="88e57-417">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="88e57-418">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-418">Requirements</span></span>

|<span data-ttu-id="88e57-419">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-419">Requirement</span></span>| <span data-ttu-id="88e57-420">値</span><span class="sxs-lookup"><span data-stu-id="88e57-420">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-421">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-421">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-422">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-422">1.0</span></span>|
|[<span data-ttu-id="88e57-423">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-423">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-424">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-424">ReadItem</span></span>|
|[<span data-ttu-id="88e57-425">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-425">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-426">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-426">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-427">例</span><span class="sxs-lookup"><span data-stu-id="88e57-427">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="88e57-428">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="88e57-428">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="88e57-429">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="88e57-429">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="88e57-430">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="88e57-430">Read mode</span></span>

<span data-ttu-id="88e57-431">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-431">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="88e57-432">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="88e57-432">Compose mode</span></span>

<span data-ttu-id="88e57-433">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-433">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-434">型:</span><span class="sxs-lookup"><span data-stu-id="88e57-434">Type:</span></span>

*   <span data-ttu-id="88e57-435">String | [Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="88e57-435">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="88e57-436">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-436">Requirements</span></span>

|<span data-ttu-id="88e57-437">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-437">Requirement</span></span>| <span data-ttu-id="88e57-438">値</span><span class="sxs-lookup"><span data-stu-id="88e57-438">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-439">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-439">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-440">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-440">1.0</span></span>|
|[<span data-ttu-id="88e57-441">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-441">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-442">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-442">ReadItem</span></span>|
|[<span data-ttu-id="88e57-443">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-443">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-444">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-444">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-445">例</span><span class="sxs-lookup"><span data-stu-id="88e57-445">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="88e57-446">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="88e57-446">normalizedSubject :String</span></span>

<span data-ttu-id="88e57-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="88e57-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="88e57-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="88e57-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-451">型:</span><span class="sxs-lookup"><span data-stu-id="88e57-451">Type:</span></span>

*   <span data-ttu-id="88e57-452">String</span><span class="sxs-lookup"><span data-stu-id="88e57-452">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="88e57-453">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-453">Requirements</span></span>

|<span data-ttu-id="88e57-454">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-454">Requirement</span></span>| <span data-ttu-id="88e57-455">値</span><span class="sxs-lookup"><span data-stu-id="88e57-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-456">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-457">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-457">1.0</span></span>|
|[<span data-ttu-id="88e57-458">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-458">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-459">ReadItem</span></span>|
|[<span data-ttu-id="88e57-460">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-460">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-461">読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-461">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-462">例</span><span class="sxs-lookup"><span data-stu-id="88e57-462">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="88e57-463">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="88e57-463">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="88e57-464">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="88e57-464">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-465">型:</span><span class="sxs-lookup"><span data-stu-id="88e57-465">Type:</span></span>

*   [<span data-ttu-id="88e57-466">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="88e57-466">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="88e57-467">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-467">Requirements</span></span>

|<span data-ttu-id="88e57-468">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-468">Requirement</span></span>| <span data-ttu-id="88e57-469">値</span><span class="sxs-lookup"><span data-stu-id="88e57-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-470">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-471">1.3</span><span class="sxs-lookup"><span data-stu-id="88e57-471">1.3</span></span>|
|[<span data-ttu-id="88e57-472">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-472">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-473">ReadItem</span></span>|
|[<span data-ttu-id="88e57-474">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-474">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-475">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-475">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="88e57-476">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="88e57-476">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="88e57-477">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="88e57-477">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="88e57-478">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="88e57-478">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="88e57-479">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="88e57-479">Read mode</span></span>

<span data-ttu-id="88e57-480">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-480">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="88e57-481">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="88e57-481">Compose mode</span></span>

<span data-ttu-id="88e57-482">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-482">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-483">種類:</span><span class="sxs-lookup"><span data-stu-id="88e57-483">Type:</span></span>

*   <span data-ttu-id="88e57-484">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="88e57-484">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="88e57-485">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-485">Requirements</span></span>

|<span data-ttu-id="88e57-486">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-486">Requirement</span></span>| <span data-ttu-id="88e57-487">値</span><span class="sxs-lookup"><span data-stu-id="88e57-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-488">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-489">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-489">1.0</span></span>|
|[<span data-ttu-id="88e57-490">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-491">ReadItem</span></span>|
|[<span data-ttu-id="88e57-492">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-493">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-493">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-494">例</span><span class="sxs-lookup"><span data-stu-id="88e57-494">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="88e57-495">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="88e57-495">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="88e57-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="88e57-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-498">型:</span><span class="sxs-lookup"><span data-stu-id="88e57-498">Type:</span></span>

*   [<span data-ttu-id="88e57-499">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="88e57-499">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="88e57-500">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-500">Requirements</span></span>

|<span data-ttu-id="88e57-501">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-501">Requirement</span></span>| <span data-ttu-id="88e57-502">値</span><span class="sxs-lookup"><span data-stu-id="88e57-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-503">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-504">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-504">1.0</span></span>|
|[<span data-ttu-id="88e57-505">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-505">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-506">ReadItem</span></span>|
|[<span data-ttu-id="88e57-507">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-507">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-508">読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-508">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-509">例</span><span class="sxs-lookup"><span data-stu-id="88e57-509">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="88e57-510">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="88e57-510">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="88e57-511">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="88e57-511">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="88e57-512">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="88e57-512">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="88e57-513">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="88e57-513">Read mode</span></span>

<span data-ttu-id="88e57-514">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-514">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="88e57-515">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="88e57-515">Compose mode</span></span>

<span data-ttu-id="88e57-516">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-516">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-517">種類:</span><span class="sxs-lookup"><span data-stu-id="88e57-517">Type:</span></span>

*   <span data-ttu-id="88e57-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="88e57-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="88e57-519">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-519">Requirements</span></span>

|<span data-ttu-id="88e57-520">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-520">Requirement</span></span>| <span data-ttu-id="88e57-521">値</span><span class="sxs-lookup"><span data-stu-id="88e57-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-522">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-523">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-523">1.0</span></span>|
|[<span data-ttu-id="88e57-524">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-525">ReadItem</span></span>|
|[<span data-ttu-id="88e57-526">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-526">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-527">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-527">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-528">例</span><span class="sxs-lookup"><span data-stu-id="88e57-528">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="88e57-529">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="88e57-529">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="88e57-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="88e57-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="88e57-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="88e57-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="88e57-534">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="88e57-534">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-535">種類:</span><span class="sxs-lookup"><span data-stu-id="88e57-535">Type:</span></span>

*   [<span data-ttu-id="88e57-536">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="88e57-536">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="88e57-537">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-537">Requirements</span></span>

|<span data-ttu-id="88e57-538">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-538">Requirement</span></span>| <span data-ttu-id="88e57-539">値</span><span class="sxs-lookup"><span data-stu-id="88e57-539">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-540">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-540">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-541">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-541">1.0</span></span>|
|[<span data-ttu-id="88e57-542">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-542">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-543">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-543">ReadItem</span></span>|
|[<span data-ttu-id="88e57-544">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-544">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-545">読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-545">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-546">例</span><span class="sxs-lookup"><span data-stu-id="88e57-546">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="88e57-547">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="88e57-547">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="88e57-548">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="88e57-548">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="88e57-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="88e57-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="88e57-551">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="88e57-551">Read mode</span></span>

<span data-ttu-id="88e57-552">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-552">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="88e57-553">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="88e57-553">Compose mode</span></span>

<span data-ttu-id="88e57-554">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-554">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="88e57-555">[`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="88e57-555">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-556">型:</span><span class="sxs-lookup"><span data-stu-id="88e57-556">Type:</span></span>

*   <span data-ttu-id="88e57-557">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="88e57-557">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="88e57-558">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-558">Requirements</span></span>

|<span data-ttu-id="88e57-559">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-559">Requirement</span></span>| <span data-ttu-id="88e57-560">値</span><span class="sxs-lookup"><span data-stu-id="88e57-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-561">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-562">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-562">1.0</span></span>|
|[<span data-ttu-id="88e57-563">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-563">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-564">ReadItem</span></span>|
|[<span data-ttu-id="88e57-565">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-565">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-566">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-566">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-567">例</span><span class="sxs-lookup"><span data-stu-id="88e57-567">Example</span></span>

<span data-ttu-id="88e57-568">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="88e57-568">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="88e57-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="88e57-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="88e57-570">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="88e57-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="88e57-571">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="88e57-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="88e57-572">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="88e57-572">Read mode</span></span>

<span data-ttu-id="88e57-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="88e57-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="88e57-575">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="88e57-575">Compose mode</span></span>

<span data-ttu-id="88e57-576">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="88e57-577">型:</span><span class="sxs-lookup"><span data-stu-id="88e57-577">Type:</span></span>

*   <span data-ttu-id="88e57-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="88e57-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="88e57-579">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-579">Requirements</span></span>

|<span data-ttu-id="88e57-580">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-580">Requirement</span></span>| <span data-ttu-id="88e57-581">値</span><span class="sxs-lookup"><span data-stu-id="88e57-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-582">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-583">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-583">1.0</span></span>|
|[<span data-ttu-id="88e57-584">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-584">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-585">ReadItem</span></span>|
|[<span data-ttu-id="88e57-586">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-586">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-587">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-587">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="88e57-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="88e57-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="88e57-589">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="88e57-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="88e57-590">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="88e57-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="88e57-591">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="88e57-591">Read mode</span></span>

<span data-ttu-id="88e57-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="88e57-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="88e57-594">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="88e57-594">Compose mode</span></span>

<span data-ttu-id="88e57-595">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="88e57-596">種類:</span><span class="sxs-lookup"><span data-stu-id="88e57-596">Type:</span></span>

*   <span data-ttu-id="88e57-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="88e57-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="88e57-598">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-598">Requirements</span></span>

|<span data-ttu-id="88e57-599">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-599">Requirement</span></span>| <span data-ttu-id="88e57-600">値</span><span class="sxs-lookup"><span data-stu-id="88e57-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-601">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-602">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-602">1.0</span></span>|
|[<span data-ttu-id="88e57-603">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-603">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-604">ReadItem</span></span>|
|[<span data-ttu-id="88e57-605">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-605">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-606">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-606">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-607">例</span><span class="sxs-lookup"><span data-stu-id="88e57-607">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="88e57-608">メソッド</span><span class="sxs-lookup"><span data-stu-id="88e57-608">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="88e57-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="88e57-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="88e57-610">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="88e57-610">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="88e57-611">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="88e57-611">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="88e57-612">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="88e57-612">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="88e57-613">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="88e57-613">Parameters:</span></span>

|<span data-ttu-id="88e57-614">名前</span><span class="sxs-lookup"><span data-stu-id="88e57-614">Name</span></span>| <span data-ttu-id="88e57-615">型</span><span class="sxs-lookup"><span data-stu-id="88e57-615">Type</span></span>| <span data-ttu-id="88e57-616">属性</span><span class="sxs-lookup"><span data-stu-id="88e57-616">Attributes</span></span>| <span data-ttu-id="88e57-617">説明</span><span class="sxs-lookup"><span data-stu-id="88e57-617">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="88e57-618">String</span><span class="sxs-lookup"><span data-stu-id="88e57-618">String</span></span>||<span data-ttu-id="88e57-p132">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="88e57-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="88e57-621">String</span><span class="sxs-lookup"><span data-stu-id="88e57-621">String</span></span>||<span data-ttu-id="88e57-p133">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="88e57-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="88e57-624">Object</span><span class="sxs-lookup"><span data-stu-id="88e57-624">Object</span></span>| <span data-ttu-id="88e57-625">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-625">&lt;optional&gt;</span></span>|<span data-ttu-id="88e57-626">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="88e57-626">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="88e57-627">Object</span><span class="sxs-lookup"><span data-stu-id="88e57-627">Object</span></span> | <span data-ttu-id="88e57-628">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-628">&lt;optional&gt;</span></span> | <span data-ttu-id="88e57-629">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="88e57-629">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="88e57-630">Boolean</span><span class="sxs-lookup"><span data-stu-id="88e57-630">Boolean</span></span> | <span data-ttu-id="88e57-631">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-631">&lt;optional&gt;</span></span> | <span data-ttu-id="88e57-632">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="88e57-632">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="88e57-633">function</span><span class="sxs-lookup"><span data-stu-id="88e57-633">function</span></span>| <span data-ttu-id="88e57-634">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-634">&lt;optional&gt;</span></span>|<span data-ttu-id="88e57-635">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-635">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="88e57-636">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-636">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="88e57-637">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="88e57-637">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="88e57-638">エラー</span><span class="sxs-lookup"><span data-stu-id="88e57-638">Errors</span></span>

| <span data-ttu-id="88e57-639">エラー コード</span><span class="sxs-lookup"><span data-stu-id="88e57-639">Error code</span></span> | <span data-ttu-id="88e57-640">説明</span><span class="sxs-lookup"><span data-stu-id="88e57-640">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="88e57-641">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="88e57-641">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="88e57-642">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="88e57-642">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="88e57-643">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="88e57-643">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="88e57-644">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-644">Requirements</span></span>

|<span data-ttu-id="88e57-645">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-645">Requirement</span></span>| <span data-ttu-id="88e57-646">値</span><span class="sxs-lookup"><span data-stu-id="88e57-646">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-647">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-647">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-648">1.1</span><span class="sxs-lookup"><span data-stu-id="88e57-648">1.1</span></span>|
|[<span data-ttu-id="88e57-649">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-649">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-650">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="88e57-650">ReadWriteItem</span></span>|
|[<span data-ttu-id="88e57-651">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-651">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-652">作成</span><span class="sxs-lookup"><span data-stu-id="88e57-652">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="88e57-653">例</span><span class="sxs-lookup"><span data-stu-id="88e57-653">Examples</span></span>

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

<span data-ttu-id="88e57-654">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="88e57-654">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="88e57-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="88e57-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="88e57-656">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="88e57-656">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="88e57-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="88e57-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="88e57-660">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="88e57-660">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="88e57-661">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="88e57-661">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="88e57-662">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="88e57-662">Parameters:</span></span>

|<span data-ttu-id="88e57-663">名前</span><span class="sxs-lookup"><span data-stu-id="88e57-663">Name</span></span>| <span data-ttu-id="88e57-664">型</span><span class="sxs-lookup"><span data-stu-id="88e57-664">Type</span></span>| <span data-ttu-id="88e57-665">属性</span><span class="sxs-lookup"><span data-stu-id="88e57-665">Attributes</span></span>| <span data-ttu-id="88e57-666">説明</span><span class="sxs-lookup"><span data-stu-id="88e57-666">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="88e57-667">String</span><span class="sxs-lookup"><span data-stu-id="88e57-667">String</span></span>||<span data-ttu-id="88e57-p135">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="88e57-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="88e57-670">String</span><span class="sxs-lookup"><span data-stu-id="88e57-670">String</span></span>||<span data-ttu-id="88e57-p136">添付するアイテムの件名。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="88e57-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="88e57-673">Object</span><span class="sxs-lookup"><span data-stu-id="88e57-673">Object</span></span>| <span data-ttu-id="88e57-674">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-674">&lt;optional&gt;</span></span>|<span data-ttu-id="88e57-675">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="88e57-675">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="88e57-676">Object</span><span class="sxs-lookup"><span data-stu-id="88e57-676">Object</span></span>| <span data-ttu-id="88e57-677">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-677">&lt;optional&gt;</span></span>|<span data-ttu-id="88e57-678">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="88e57-678">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="88e57-679">function</span><span class="sxs-lookup"><span data-stu-id="88e57-679">function</span></span>| <span data-ttu-id="88e57-680">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-680">&lt;optional&gt;</span></span>|<span data-ttu-id="88e57-681">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-681">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="88e57-682">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-682">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="88e57-683">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="88e57-683">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="88e57-684">エラー</span><span class="sxs-lookup"><span data-stu-id="88e57-684">Errors</span></span>

| <span data-ttu-id="88e57-685">エラー コード</span><span class="sxs-lookup"><span data-stu-id="88e57-685">Error code</span></span> | <span data-ttu-id="88e57-686">説明</span><span class="sxs-lookup"><span data-stu-id="88e57-686">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="88e57-687">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="88e57-687">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="88e57-688">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-688">Requirements</span></span>

|<span data-ttu-id="88e57-689">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-689">Requirement</span></span>| <span data-ttu-id="88e57-690">値</span><span class="sxs-lookup"><span data-stu-id="88e57-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-691">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-692">1.1</span><span class="sxs-lookup"><span data-stu-id="88e57-692">1.1</span></span>|
|[<span data-ttu-id="88e57-693">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-693">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-694">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="88e57-694">ReadWriteItem</span></span>|
|[<span data-ttu-id="88e57-695">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-695">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-696">作成</span><span class="sxs-lookup"><span data-stu-id="88e57-696">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-697">例</span><span class="sxs-lookup"><span data-stu-id="88e57-697">Example</span></span>

<span data-ttu-id="88e57-698">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-698">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="88e57-699">close()</span><span class="sxs-lookup"><span data-stu-id="88e57-699">close()</span></span>

<span data-ttu-id="88e57-700">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="88e57-700">Closes the current item that is being composed.</span></span>

<span data-ttu-id="88e57-p137">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="88e57-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="88e57-703">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-703">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="88e57-704">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="88e57-704">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="88e57-705">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-705">Requirements</span></span>

|<span data-ttu-id="88e57-706">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-706">Requirement</span></span>| <span data-ttu-id="88e57-707">値</span><span class="sxs-lookup"><span data-stu-id="88e57-707">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-708">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-708">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-709">1.3</span><span class="sxs-lookup"><span data-stu-id="88e57-709">1.3</span></span>|
|[<span data-ttu-id="88e57-710">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-710">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-711">制限あり</span><span class="sxs-lookup"><span data-stu-id="88e57-711">Restricted</span></span>|
|[<span data-ttu-id="88e57-712">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-712">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-713">作成</span><span class="sxs-lookup"><span data-stu-id="88e57-713">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="88e57-714">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="88e57-714">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="88e57-715">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-715">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="88e57-716">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="88e57-716">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="88e57-717">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-717">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="88e57-718">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="88e57-718">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="88e57-p138">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="88e57-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="88e57-722">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="88e57-722">Parameters:</span></span>

| <span data-ttu-id="88e57-723">名前</span><span class="sxs-lookup"><span data-stu-id="88e57-723">Name</span></span> | <span data-ttu-id="88e57-724">型</span><span class="sxs-lookup"><span data-stu-id="88e57-724">Type</span></span> | <span data-ttu-id="88e57-725">属性</span><span class="sxs-lookup"><span data-stu-id="88e57-725">Attributes</span></span> | <span data-ttu-id="88e57-726">説明</span><span class="sxs-lookup"><span data-stu-id="88e57-726">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="88e57-727">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="88e57-727">String &#124; Object</span></span>| |<span data-ttu-id="88e57-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="88e57-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="88e57-730">**または**</span><span class="sxs-lookup"><span data-stu-id="88e57-730">**OR**</span></span><br/><span data-ttu-id="88e57-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="88e57-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="88e57-733">String</span><span class="sxs-lookup"><span data-stu-id="88e57-733">String</span></span> | <span data-ttu-id="88e57-734">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-734">&lt;optional&gt;</span></span> | <span data-ttu-id="88e57-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="88e57-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="88e57-737">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-737">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="88e57-738">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-738">&lt;optional&gt;</span></span> | <span data-ttu-id="88e57-739">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="88e57-739">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="88e57-740">String</span><span class="sxs-lookup"><span data-stu-id="88e57-740">String</span></span> | | <span data-ttu-id="88e57-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="88e57-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="88e57-743">String</span><span class="sxs-lookup"><span data-stu-id="88e57-743">String</span></span> | | <span data-ttu-id="88e57-744">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="88e57-744">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="88e57-745">String</span><span class="sxs-lookup"><span data-stu-id="88e57-745">String</span></span> | | <span data-ttu-id="88e57-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="88e57-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="88e57-748">ブール値</span><span class="sxs-lookup"><span data-stu-id="88e57-748">Boolean</span></span> | | <span data-ttu-id="88e57-p144">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="88e57-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="88e57-751">String</span><span class="sxs-lookup"><span data-stu-id="88e57-751">String</span></span> | | <span data-ttu-id="88e57-p145">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="88e57-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="88e57-755">function</span><span class="sxs-lookup"><span data-stu-id="88e57-755">function</span></span> | <span data-ttu-id="88e57-756">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-756">&lt;optional&gt;</span></span> | <span data-ttu-id="88e57-757">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-757">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="88e57-758">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-758">Requirements</span></span>

|<span data-ttu-id="88e57-759">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-759">Requirement</span></span>| <span data-ttu-id="88e57-760">値</span><span class="sxs-lookup"><span data-stu-id="88e57-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-761">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-761">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-762">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-762">1.0</span></span>|
|[<span data-ttu-id="88e57-763">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-763">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-764">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-764">ReadItem</span></span>|
|[<span data-ttu-id="88e57-765">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-765">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-766">読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-766">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="88e57-767">例</span><span class="sxs-lookup"><span data-stu-id="88e57-767">Examples</span></span>

<span data-ttu-id="88e57-768">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="88e57-768">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="88e57-769">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="88e57-769">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="88e57-770">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="88e57-770">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="88e57-771">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="88e57-771">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="88e57-772">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="88e57-772">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="88e57-773">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="88e57-773">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="88e57-774">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="88e57-774">displayReplyForm(formData)</span></span>

<span data-ttu-id="88e57-775">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-775">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="88e57-776">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="88e57-776">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="88e57-777">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-777">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="88e57-778">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="88e57-778">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="88e57-p146">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="88e57-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="88e57-782">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="88e57-782">Parameters:</span></span>

| <span data-ttu-id="88e57-783">名前</span><span class="sxs-lookup"><span data-stu-id="88e57-783">Name</span></span> | <span data-ttu-id="88e57-784">型</span><span class="sxs-lookup"><span data-stu-id="88e57-784">Type</span></span> | <span data-ttu-id="88e57-785">属性</span><span class="sxs-lookup"><span data-stu-id="88e57-785">Attributes</span></span> | <span data-ttu-id="88e57-786">説明</span><span class="sxs-lookup"><span data-stu-id="88e57-786">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="88e57-787">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="88e57-787">String &#124; Object</span></span>| | <span data-ttu-id="88e57-p147">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="88e57-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="88e57-790">**または**</span><span class="sxs-lookup"><span data-stu-id="88e57-790">**OR**</span></span><br/><span data-ttu-id="88e57-p148">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="88e57-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="88e57-793">String</span><span class="sxs-lookup"><span data-stu-id="88e57-793">String</span></span> | <span data-ttu-id="88e57-794">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-794">&lt;optional&gt;</span></span> | <span data-ttu-id="88e57-p149">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="88e57-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="88e57-797">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-797">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="88e57-798">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-798">&lt;optional&gt;</span></span> | <span data-ttu-id="88e57-799">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="88e57-799">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="88e57-800">String</span><span class="sxs-lookup"><span data-stu-id="88e57-800">String</span></span> | | <span data-ttu-id="88e57-p150">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="88e57-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="88e57-803">String</span><span class="sxs-lookup"><span data-stu-id="88e57-803">String</span></span> | | <span data-ttu-id="88e57-804">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="88e57-804">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="88e57-805">String</span><span class="sxs-lookup"><span data-stu-id="88e57-805">String</span></span> | | <span data-ttu-id="88e57-p151">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="88e57-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="88e57-808">ブール値</span><span class="sxs-lookup"><span data-stu-id="88e57-808">Boolean</span></span> | | <span data-ttu-id="88e57-p152">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="88e57-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="88e57-811">String</span><span class="sxs-lookup"><span data-stu-id="88e57-811">String</span></span> | | <span data-ttu-id="88e57-p153">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="88e57-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="88e57-815">function</span><span class="sxs-lookup"><span data-stu-id="88e57-815">function</span></span> | <span data-ttu-id="88e57-816">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-816">&lt;optional&gt;</span></span> | <span data-ttu-id="88e57-817">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-817">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="88e57-818">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-818">Requirements</span></span>

|<span data-ttu-id="88e57-819">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-819">Requirement</span></span>| <span data-ttu-id="88e57-820">値</span><span class="sxs-lookup"><span data-stu-id="88e57-820">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-821">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-821">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-822">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-822">1.0</span></span>|
|[<span data-ttu-id="88e57-823">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-823">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-824">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-824">ReadItem</span></span>|
|[<span data-ttu-id="88e57-825">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-825">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-826">読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-826">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="88e57-827">例</span><span class="sxs-lookup"><span data-stu-id="88e57-827">Examples</span></span>

<span data-ttu-id="88e57-828">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="88e57-828">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="88e57-829">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="88e57-829">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="88e57-830">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="88e57-830">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="88e57-831">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="88e57-831">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="88e57-832">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="88e57-832">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="88e57-833">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="88e57-833">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="88e57-834">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="88e57-834">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="88e57-835">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="88e57-835">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="88e57-836">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="88e57-836">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="88e57-837">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-837">Requirements</span></span>

|<span data-ttu-id="88e57-838">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-838">Requirement</span></span>| <span data-ttu-id="88e57-839">値</span><span class="sxs-lookup"><span data-stu-id="88e57-839">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-840">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-840">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-841">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-841">1.0</span></span>|
|[<span data-ttu-id="88e57-842">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-842">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-843">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-843">ReadItem</span></span>|
|[<span data-ttu-id="88e57-844">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-844">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-845">読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-845">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="88e57-846">戻り値:</span><span class="sxs-lookup"><span data-stu-id="88e57-846">Returns:</span></span>

<span data-ttu-id="88e57-847">型:[Entities](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="88e57-847">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="88e57-848">例</span><span class="sxs-lookup"><span data-stu-id="88e57-848">Example</span></span>

<span data-ttu-id="88e57-849">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="88e57-849">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="88e57-850">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="88e57-850">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="88e57-851">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="88e57-851">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="88e57-852">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="88e57-852">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="88e57-853">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="88e57-853">Parameters:</span></span>

|<span data-ttu-id="88e57-854">名前</span><span class="sxs-lookup"><span data-stu-id="88e57-854">Name</span></span>| <span data-ttu-id="88e57-855">型</span><span class="sxs-lookup"><span data-stu-id="88e57-855">Type</span></span>| <span data-ttu-id="88e57-856">説明</span><span class="sxs-lookup"><span data-stu-id="88e57-856">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="88e57-857">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="88e57-857">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="88e57-858">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="88e57-858">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="88e57-859">Requirements</span><span class="sxs-lookup"><span data-stu-id="88e57-859">Requirements</span></span>

|<span data-ttu-id="88e57-860">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-860">Requirement</span></span>| <span data-ttu-id="88e57-861">値</span><span class="sxs-lookup"><span data-stu-id="88e57-861">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-862">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-862">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-863">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-863">1.0</span></span>|
|[<span data-ttu-id="88e57-864">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-864">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-865">制限あり</span><span class="sxs-lookup"><span data-stu-id="88e57-865">Restricted</span></span>|
|[<span data-ttu-id="88e57-866">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-866">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-867">読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-867">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="88e57-868">戻り値:</span><span class="sxs-lookup"><span data-stu-id="88e57-868">Returns:</span></span>

<span data-ttu-id="88e57-869">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-869">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="88e57-870">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-870">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="88e57-871">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="88e57-871">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="88e57-872">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="88e57-872">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="88e57-873">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="88e57-873">Value of `entityType`</span></span> | <span data-ttu-id="88e57-874">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="88e57-874">Type of objects in returned array</span></span> | <span data-ttu-id="88e57-875">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="88e57-875">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="88e57-876">文字列</span><span class="sxs-lookup"><span data-stu-id="88e57-876">String</span></span> | <span data-ttu-id="88e57-877">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="88e57-877">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="88e57-878">連絡先</span><span class="sxs-lookup"><span data-stu-id="88e57-878">Contact</span></span> | <span data-ttu-id="88e57-879">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="88e57-879">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="88e57-880">文字列</span><span class="sxs-lookup"><span data-stu-id="88e57-880">String</span></span> | <span data-ttu-id="88e57-881">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="88e57-881">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="88e57-882">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="88e57-882">MeetingSuggestion</span></span> | <span data-ttu-id="88e57-883">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="88e57-883">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="88e57-884">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="88e57-884">PhoneNumber</span></span> | <span data-ttu-id="88e57-885">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="88e57-885">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="88e57-886">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="88e57-886">TaskSuggestion</span></span> | <span data-ttu-id="88e57-887">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="88e57-887">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="88e57-888">文字列</span><span class="sxs-lookup"><span data-stu-id="88e57-888">String</span></span> | <span data-ttu-id="88e57-889">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="88e57-889">**Restricted**</span></span> |

<span data-ttu-id="88e57-890">型:Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="88e57-890">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="88e57-891">例</span><span class="sxs-lookup"><span data-stu-id="88e57-891">Example</span></span>

<span data-ttu-id="88e57-892">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="88e57-892">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="88e57-893">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="88e57-893">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="88e57-894">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-894">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="88e57-895">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="88e57-895">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="88e57-896">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-896">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="88e57-897">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="88e57-897">Parameters:</span></span>

|<span data-ttu-id="88e57-898">名前</span><span class="sxs-lookup"><span data-stu-id="88e57-898">Name</span></span>| <span data-ttu-id="88e57-899">型</span><span class="sxs-lookup"><span data-stu-id="88e57-899">Type</span></span>| <span data-ttu-id="88e57-900">説明</span><span class="sxs-lookup"><span data-stu-id="88e57-900">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="88e57-901">String</span><span class="sxs-lookup"><span data-stu-id="88e57-901">String</span></span>|<span data-ttu-id="88e57-902">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="88e57-902">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="88e57-903">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-903">Requirements</span></span>

|<span data-ttu-id="88e57-904">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-904">Requirement</span></span>| <span data-ttu-id="88e57-905">値</span><span class="sxs-lookup"><span data-stu-id="88e57-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-906">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-907">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-907">1.0</span></span>|
|[<span data-ttu-id="88e57-908">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-909">ReadItem</span></span>|
|[<span data-ttu-id="88e57-910">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-911">読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-911">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="88e57-912">戻り値:</span><span class="sxs-lookup"><span data-stu-id="88e57-912">Returns:</span></span>

<span data-ttu-id="88e57-p155">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="88e57-915">型:Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="88e57-915">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="88e57-916">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="88e57-916">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="88e57-917">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-917">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="88e57-918">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="88e57-918">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="88e57-p156">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="88e57-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="88e57-922">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="88e57-922">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="88e57-923">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="88e57-923">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="88e57-p157">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="88e57-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="88e57-927">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-927">Requirements</span></span>

|<span data-ttu-id="88e57-928">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-928">Requirement</span></span>| <span data-ttu-id="88e57-929">値</span><span class="sxs-lookup"><span data-stu-id="88e57-929">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-930">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-930">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-931">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-931">1.0</span></span>|
|[<span data-ttu-id="88e57-932">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-932">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-933">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-933">ReadItem</span></span>|
|[<span data-ttu-id="88e57-934">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-934">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-935">読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-935">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="88e57-936">戻り値:</span><span class="sxs-lookup"><span data-stu-id="88e57-936">Returns:</span></span>

<span data-ttu-id="88e57-p158">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="88e57-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="88e57-939">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="88e57-939">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="88e57-940">Object</span><span class="sxs-lookup"><span data-stu-id="88e57-940">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="88e57-941">例</span><span class="sxs-lookup"><span data-stu-id="88e57-941">Example</span></span>

<span data-ttu-id="88e57-942">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="88e57-942">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="88e57-943">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="88e57-943">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="88e57-944">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-944">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="88e57-945">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="88e57-945">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="88e57-946">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="88e57-946">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="88e57-p159">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="88e57-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="88e57-949">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="88e57-949">Parameters:</span></span>

|<span data-ttu-id="88e57-950">名前</span><span class="sxs-lookup"><span data-stu-id="88e57-950">Name</span></span>| <span data-ttu-id="88e57-951">型</span><span class="sxs-lookup"><span data-stu-id="88e57-951">Type</span></span>| <span data-ttu-id="88e57-952">説明</span><span class="sxs-lookup"><span data-stu-id="88e57-952">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="88e57-953">String</span><span class="sxs-lookup"><span data-stu-id="88e57-953">String</span></span>|<span data-ttu-id="88e57-954">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="88e57-954">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="88e57-955">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-955">Requirements</span></span>

|<span data-ttu-id="88e57-956">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-956">Requirement</span></span>| <span data-ttu-id="88e57-957">値</span><span class="sxs-lookup"><span data-stu-id="88e57-957">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-958">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-958">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-959">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-959">1.0</span></span>|
|[<span data-ttu-id="88e57-960">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-960">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-961">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-961">ReadItem</span></span>|
|[<span data-ttu-id="88e57-962">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-962">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-963">読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-963">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="88e57-964">戻り値:</span><span class="sxs-lookup"><span data-stu-id="88e57-964">Returns:</span></span>

<span data-ttu-id="88e57-965">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="88e57-965">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="88e57-966">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="88e57-966">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="88e57-967">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="88e57-967">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="88e57-968">例</span><span class="sxs-lookup"><span data-stu-id="88e57-968">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="88e57-969">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="88e57-969">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="88e57-970">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-970">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="88e57-p160">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="88e57-973">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="88e57-973">Parameters:</span></span>

|<span data-ttu-id="88e57-974">名前</span><span class="sxs-lookup"><span data-stu-id="88e57-974">Name</span></span>| <span data-ttu-id="88e57-975">型</span><span class="sxs-lookup"><span data-stu-id="88e57-975">Type</span></span>| <span data-ttu-id="88e57-976">属性</span><span class="sxs-lookup"><span data-stu-id="88e57-976">Attributes</span></span>| <span data-ttu-id="88e57-977">説明</span><span class="sxs-lookup"><span data-stu-id="88e57-977">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="88e57-978">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="88e57-978">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="88e57-p161">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="88e57-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="88e57-982">Object</span><span class="sxs-lookup"><span data-stu-id="88e57-982">Object</span></span>| <span data-ttu-id="88e57-983">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-983">&lt;optional&gt;</span></span>|<span data-ttu-id="88e57-984">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="88e57-984">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="88e57-985">Object</span><span class="sxs-lookup"><span data-stu-id="88e57-985">Object</span></span>| <span data-ttu-id="88e57-986">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-986">&lt;optional&gt;</span></span>|<span data-ttu-id="88e57-987">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="88e57-987">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="88e57-988">function</span><span class="sxs-lookup"><span data-stu-id="88e57-988">function</span></span>||<span data-ttu-id="88e57-989">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-989">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="88e57-990">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="88e57-990">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="88e57-991">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="88e57-991">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="88e57-992">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-992">Requirements</span></span>

|<span data-ttu-id="88e57-993">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-993">Requirement</span></span>| <span data-ttu-id="88e57-994">値</span><span class="sxs-lookup"><span data-stu-id="88e57-994">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-995">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-995">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-996">1.2</span><span class="sxs-lookup"><span data-stu-id="88e57-996">1.2</span></span>|
|[<span data-ttu-id="88e57-997">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-997">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-998">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="88e57-998">ReadWriteItem</span></span>|
|[<span data-ttu-id="88e57-999">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-999">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-1000">作成</span><span class="sxs-lookup"><span data-stu-id="88e57-1000">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="88e57-1001">戻り値:</span><span class="sxs-lookup"><span data-stu-id="88e57-1001">Returns:</span></span>

<span data-ttu-id="88e57-1002">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="88e57-1002">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="88e57-1003">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="88e57-1003">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="88e57-1004">String</span><span class="sxs-lookup"><span data-stu-id="88e57-1004">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="88e57-1005">例</span><span class="sxs-lookup"><span data-stu-id="88e57-1005">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="88e57-1006">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="88e57-1006">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="88e57-1007">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="88e57-1007">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="88e57-p163">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="88e57-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="88e57-1011">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="88e57-1011">Parameters:</span></span>

|<span data-ttu-id="88e57-1012">名前</span><span class="sxs-lookup"><span data-stu-id="88e57-1012">Name</span></span>| <span data-ttu-id="88e57-1013">型</span><span class="sxs-lookup"><span data-stu-id="88e57-1013">Type</span></span>| <span data-ttu-id="88e57-1014">属性</span><span class="sxs-lookup"><span data-stu-id="88e57-1014">Attributes</span></span>| <span data-ttu-id="88e57-1015">説明</span><span class="sxs-lookup"><span data-stu-id="88e57-1015">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="88e57-1016">function</span><span class="sxs-lookup"><span data-stu-id="88e57-1016">function</span></span>||<span data-ttu-id="88e57-1017">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-1017">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="88e57-1018">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-1018">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="88e57-1019">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="88e57-1019">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="88e57-1020">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="88e57-1020">Object</span></span>| <span data-ttu-id="88e57-1021">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-1021">&lt;optional&gt;</span></span>|<span data-ttu-id="88e57-1022">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="88e57-1022">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="88e57-1023">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="88e57-1023">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="88e57-1024">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-1024">Requirements</span></span>

|<span data-ttu-id="88e57-1025">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-1025">Requirement</span></span>| <span data-ttu-id="88e57-1026">値</span><span class="sxs-lookup"><span data-stu-id="88e57-1026">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-1027">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-1027">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-1028">1.0</span><span class="sxs-lookup"><span data-stu-id="88e57-1028">1.0</span></span>|
|[<span data-ttu-id="88e57-1029">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-1029">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-1030">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88e57-1030">ReadItem</span></span>|
|[<span data-ttu-id="88e57-1031">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-1031">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-1032">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="88e57-1032">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-1033">例</span><span class="sxs-lookup"><span data-stu-id="88e57-1033">Example</span></span>

<span data-ttu-id="88e57-p166">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="88e57-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="88e57-1037">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="88e57-1037">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="88e57-1038">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="88e57-1038">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="88e57-p167">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="88e57-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="88e57-1043">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="88e57-1043">Parameters:</span></span>

|<span data-ttu-id="88e57-1044">名前</span><span class="sxs-lookup"><span data-stu-id="88e57-1044">Name</span></span>| <span data-ttu-id="88e57-1045">型</span><span class="sxs-lookup"><span data-stu-id="88e57-1045">Type</span></span>| <span data-ttu-id="88e57-1046">属性</span><span class="sxs-lookup"><span data-stu-id="88e57-1046">Attributes</span></span>| <span data-ttu-id="88e57-1047">説明</span><span class="sxs-lookup"><span data-stu-id="88e57-1047">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="88e57-1048">String</span><span class="sxs-lookup"><span data-stu-id="88e57-1048">String</span></span>||<span data-ttu-id="88e57-1049">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="88e57-1049">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="88e57-1050">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="88e57-1050">Object</span></span>| <span data-ttu-id="88e57-1051">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-1051">&lt;optional&gt;</span></span>|<span data-ttu-id="88e57-1052">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="88e57-1052">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="88e57-1053">Object</span><span class="sxs-lookup"><span data-stu-id="88e57-1053">Object</span></span>| <span data-ttu-id="88e57-1054">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-1054">&lt;optional&gt;</span></span>|<span data-ttu-id="88e57-1055">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="88e57-1055">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="88e57-1056">function</span><span class="sxs-lookup"><span data-stu-id="88e57-1056">function</span></span>| <span data-ttu-id="88e57-1057">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="88e57-1058">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="88e57-1059">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="88e57-1059">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="88e57-1060">エラー</span><span class="sxs-lookup"><span data-stu-id="88e57-1060">Errors</span></span>

| <span data-ttu-id="88e57-1061">エラー コード</span><span class="sxs-lookup"><span data-stu-id="88e57-1061">Error code</span></span> | <span data-ttu-id="88e57-1062">説明</span><span class="sxs-lookup"><span data-stu-id="88e57-1062">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="88e57-1063">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="88e57-1063">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="88e57-1064">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-1064">Requirements</span></span>

|<span data-ttu-id="88e57-1065">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-1065">Requirement</span></span>| <span data-ttu-id="88e57-1066">値</span><span class="sxs-lookup"><span data-stu-id="88e57-1066">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-1067">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-1067">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-1068">1.1</span><span class="sxs-lookup"><span data-stu-id="88e57-1068">1.1</span></span>|
|[<span data-ttu-id="88e57-1069">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-1069">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-1070">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="88e57-1070">ReadWriteItem</span></span>|
|[<span data-ttu-id="88e57-1071">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-1071">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-1072">作成</span><span class="sxs-lookup"><span data-stu-id="88e57-1072">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-1073">例</span><span class="sxs-lookup"><span data-stu-id="88e57-1073">Example</span></span>

<span data-ttu-id="88e57-1074">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="88e57-1074">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="88e57-1075">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="88e57-1075">saveAsync([options], callback)</span></span>

<span data-ttu-id="88e57-1076">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="88e57-1076">Asynchronously saves an item.</span></span>

<span data-ttu-id="88e57-p168">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-p168">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="88e57-1080">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="88e57-1080">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="88e57-1081">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-1081">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="88e57-p170">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-p170">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="88e57-1085">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="88e57-1085">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="88e57-1086">Mac Outlook では、新規作成モードの会議で `saveAsync` をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="88e57-1086">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="88e57-1087">Mac Outlook では、会議で `saveAsync` を呼び出すとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-1087">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="88e57-1088">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-1088">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="88e57-1089">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="88e57-1089">Parameters:</span></span>

|<span data-ttu-id="88e57-1090">名前</span><span class="sxs-lookup"><span data-stu-id="88e57-1090">Name</span></span>| <span data-ttu-id="88e57-1091">型</span><span class="sxs-lookup"><span data-stu-id="88e57-1091">Type</span></span>| <span data-ttu-id="88e57-1092">属性</span><span class="sxs-lookup"><span data-stu-id="88e57-1092">Attributes</span></span>| <span data-ttu-id="88e57-1093">説明</span><span class="sxs-lookup"><span data-stu-id="88e57-1093">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="88e57-1094">Object</span><span class="sxs-lookup"><span data-stu-id="88e57-1094">Object</span></span>| <span data-ttu-id="88e57-1095">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="88e57-1096">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="88e57-1096">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="88e57-1097">Object</span><span class="sxs-lookup"><span data-stu-id="88e57-1097">Object</span></span>| <span data-ttu-id="88e57-1098">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="88e57-1099">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="88e57-1099">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="88e57-1100">function</span><span class="sxs-lookup"><span data-stu-id="88e57-1100">function</span></span>||<span data-ttu-id="88e57-1101">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="88e57-1102">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-1102">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="88e57-1103">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-1103">Requirements</span></span>

|<span data-ttu-id="88e57-1104">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-1104">Requirement</span></span>| <span data-ttu-id="88e57-1105">値</span><span class="sxs-lookup"><span data-stu-id="88e57-1105">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-1106">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-1106">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-1107">1.3</span><span class="sxs-lookup"><span data-stu-id="88e57-1107">1.3</span></span>|
|[<span data-ttu-id="88e57-1108">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-1108">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-1109">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="88e57-1109">ReadWriteItem</span></span>|
|[<span data-ttu-id="88e57-1110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-1110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-1111">作成</span><span class="sxs-lookup"><span data-stu-id="88e57-1111">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="88e57-1112">例</span><span class="sxs-lookup"><span data-stu-id="88e57-1112">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="88e57-p172">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="88e57-p172">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="88e57-1115">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="88e57-1115">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="88e57-1116">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="88e57-1116">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="88e57-p173">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="88e57-p173">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="88e57-1120">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="88e57-1120">Parameters:</span></span>

|<span data-ttu-id="88e57-1121">名前</span><span class="sxs-lookup"><span data-stu-id="88e57-1121">Name</span></span>| <span data-ttu-id="88e57-1122">型</span><span class="sxs-lookup"><span data-stu-id="88e57-1122">Type</span></span>| <span data-ttu-id="88e57-1123">属性</span><span class="sxs-lookup"><span data-stu-id="88e57-1123">Attributes</span></span>| <span data-ttu-id="88e57-1124">説明</span><span class="sxs-lookup"><span data-stu-id="88e57-1124">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="88e57-1125">String</span><span class="sxs-lookup"><span data-stu-id="88e57-1125">String</span></span>||<span data-ttu-id="88e57-p174">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="88e57-p174">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="88e57-1129">Object</span><span class="sxs-lookup"><span data-stu-id="88e57-1129">Object</span></span>| <span data-ttu-id="88e57-1130">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="88e57-1131">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="88e57-1131">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="88e57-1132">Object</span><span class="sxs-lookup"><span data-stu-id="88e57-1132">Object</span></span>| <span data-ttu-id="88e57-1133">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="88e57-1134">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="88e57-1134">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="88e57-1135">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="88e57-1135">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="88e57-1136">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="88e57-1136">&lt;optional&gt;</span></span>|<span data-ttu-id="88e57-p175">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-p175">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="88e57-p176">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-p176">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="88e57-1141">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-1141">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="88e57-1142">function</span><span class="sxs-lookup"><span data-stu-id="88e57-1142">function</span></span>||<span data-ttu-id="88e57-1143">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="88e57-1143">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="88e57-1144">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-1144">Requirements</span></span>

|<span data-ttu-id="88e57-1145">要件</span><span class="sxs-lookup"><span data-stu-id="88e57-1145">Requirement</span></span>| <span data-ttu-id="88e57-1146">値</span><span class="sxs-lookup"><span data-stu-id="88e57-1146">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e57-1147">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="88e57-1147">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e57-1148">1.2</span><span class="sxs-lookup"><span data-stu-id="88e57-1148">1.2</span></span>|
|[<span data-ttu-id="88e57-1149">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="88e57-1149">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88e57-1150">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="88e57-1150">ReadWriteItem</span></span>|
|[<span data-ttu-id="88e57-1151">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="88e57-1151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e57-1152">作成</span><span class="sxs-lookup"><span data-stu-id="88e57-1152">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="88e57-1153">例</span><span class="sxs-lookup"><span data-stu-id="88e57-1153">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
