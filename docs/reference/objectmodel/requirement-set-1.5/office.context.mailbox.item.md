---
title: Office.context.mailbox.item - requirement set 1.5
description: ''
ms.date: 02/15/2019
localization_priority: Priority
ms.openlocfilehash: 999a9b732a7efd3f4a9c289e5ff0b97e87278e60
ms.sourcegitcommit: 8fb60c3a31faedaea8b51b46238eb80c590a2491
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/14/2019
ms.locfileid: "30600278"
---
# <a name="item"></a><span data-ttu-id="5812c-102">item</span><span class="sxs-lookup"><span data-stu-id="5812c-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="5812c-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="5812c-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="5812c-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="5812c-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5812c-106">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-106">Requirements</span></span>

|<span data-ttu-id="5812c-107">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-107">Requirement</span></span>| <span data-ttu-id="5812c-108">値</span><span class="sxs-lookup"><span data-stu-id="5812c-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-110">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-110">1.0</span></span>|
|[<span data-ttu-id="5812c-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="5812c-112">Restricted</span></span>|
|[<span data-ttu-id="5812c-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5812c-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="5812c-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="5812c-115">Members and methods</span></span>

| <span data-ttu-id="5812c-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-116">Member</span></span> | <span data-ttu-id="5812c-117">種類</span><span class="sxs-lookup"><span data-stu-id="5812c-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="5812c-118">attachments</span><span class="sxs-lookup"><span data-stu-id="5812c-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="5812c-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-119">Member</span></span> |
| [<span data-ttu-id="5812c-120">bcc</span><span class="sxs-lookup"><span data-stu-id="5812c-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="5812c-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-121">Member</span></span> |
| [<span data-ttu-id="5812c-122">body</span><span class="sxs-lookup"><span data-stu-id="5812c-122">body</span></span>](#body-body) | <span data-ttu-id="5812c-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-123">Member</span></span> |
| [<span data-ttu-id="5812c-124">cc</span><span class="sxs-lookup"><span data-stu-id="5812c-124">cc</span></span>](#cc-arrayemailaddressdetails) | <span data-ttu-id="5812c-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-125">Member</span></span> |
| [<span data-ttu-id="5812c-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="5812c-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="5812c-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-127">Member</span></span> |
| [<span data-ttu-id="5812c-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="5812c-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="5812c-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-129">Member</span></span> |
| [<span data-ttu-id="5812c-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="5812c-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="5812c-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-131">Member</span></span> |
| [<span data-ttu-id="5812c-132">end</span><span class="sxs-lookup"><span data-stu-id="5812c-132">end</span></span>](#end-datetime) | <span data-ttu-id="5812c-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-133">Member</span></span> |
| [<span data-ttu-id="5812c-134">from</span><span class="sxs-lookup"><span data-stu-id="5812c-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="5812c-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-135">Member</span></span> |
| [<span data-ttu-id="5812c-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="5812c-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="5812c-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-137">Member</span></span> |
| [<span data-ttu-id="5812c-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="5812c-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="5812c-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-139">Member</span></span> |
| [<span data-ttu-id="5812c-140">itemId</span><span class="sxs-lookup"><span data-stu-id="5812c-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="5812c-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-141">Member</span></span> |
| [<span data-ttu-id="5812c-142">itemType</span><span class="sxs-lookup"><span data-stu-id="5812c-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="5812c-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-143">Member</span></span> |
| [<span data-ttu-id="5812c-144">location</span><span class="sxs-lookup"><span data-stu-id="5812c-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="5812c-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-145">Member</span></span> |
| [<span data-ttu-id="5812c-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="5812c-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="5812c-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-147">Member</span></span> |
| [<span data-ttu-id="5812c-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="5812c-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="5812c-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-149">Member</span></span> |
| [<span data-ttu-id="5812c-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="5812c-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetails) | <span data-ttu-id="5812c-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-151">Member</span></span> |
| [<span data-ttu-id="5812c-152">organizer</span><span class="sxs-lookup"><span data-stu-id="5812c-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="5812c-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-153">Member</span></span> |
| [<span data-ttu-id="5812c-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="5812c-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetails) | <span data-ttu-id="5812c-155">Member</span><span class="sxs-lookup"><span data-stu-id="5812c-155">Member</span></span> |
| [<span data-ttu-id="5812c-156">sender</span><span class="sxs-lookup"><span data-stu-id="5812c-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="5812c-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-157">Member</span></span> |
| [<span data-ttu-id="5812c-158">start</span><span class="sxs-lookup"><span data-stu-id="5812c-158">start</span></span>](#start-datetime) | <span data-ttu-id="5812c-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-159">Member</span></span> |
| [<span data-ttu-id="5812c-160">subject</span><span class="sxs-lookup"><span data-stu-id="5812c-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="5812c-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-161">Member</span></span> |
| [<span data-ttu-id="5812c-162">to</span><span class="sxs-lookup"><span data-stu-id="5812c-162">to</span></span>](#to-arrayemailaddressdetails) | <span data-ttu-id="5812c-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-163">Member</span></span> |
| [<span data-ttu-id="5812c-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="5812c-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="5812c-165">メソッド</span><span class="sxs-lookup"><span data-stu-id="5812c-165">Method</span></span> |
| [<span data-ttu-id="5812c-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="5812c-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="5812c-167">メソッド</span><span class="sxs-lookup"><span data-stu-id="5812c-167">Method</span></span> |
| [<span data-ttu-id="5812c-168">close</span><span class="sxs-lookup"><span data-stu-id="5812c-168">close</span></span>](#close) | <span data-ttu-id="5812c-169">メソッド</span><span class="sxs-lookup"><span data-stu-id="5812c-169">Method</span></span> |
| [<span data-ttu-id="5812c-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="5812c-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="5812c-171">メソッド</span><span class="sxs-lookup"><span data-stu-id="5812c-171">Method</span></span> |
| [<span data-ttu-id="5812c-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="5812c-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="5812c-173">メソッド</span><span class="sxs-lookup"><span data-stu-id="5812c-173">Method</span></span> |
| [<span data-ttu-id="5812c-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="5812c-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="5812c-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="5812c-175">Method</span></span> |
| [<span data-ttu-id="5812c-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="5812c-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontact) | <span data-ttu-id="5812c-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="5812c-177">Method</span></span> |
| [<span data-ttu-id="5812c-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="5812c-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontact) | <span data-ttu-id="5812c-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="5812c-179">Method</span></span> |
| [<span data-ttu-id="5812c-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="5812c-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="5812c-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="5812c-181">Method</span></span> |
| [<span data-ttu-id="5812c-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="5812c-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="5812c-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="5812c-183">Method</span></span> |
| [<span data-ttu-id="5812c-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="5812c-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="5812c-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="5812c-185">Method</span></span> |
| [<span data-ttu-id="5812c-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="5812c-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="5812c-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="5812c-187">Method</span></span> |
| [<span data-ttu-id="5812c-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="5812c-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="5812c-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="5812c-189">Method</span></span> |
| [<span data-ttu-id="5812c-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="5812c-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="5812c-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="5812c-191">Method</span></span> |
| [<span data-ttu-id="5812c-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="5812c-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="5812c-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="5812c-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="5812c-194">例</span><span class="sxs-lookup"><span data-stu-id="5812c-194">Example</span></span>

<span data-ttu-id="5812c-195">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="5812c-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
};
```

### <a name="members"></a><span data-ttu-id="5812c-196">メンバー</span><span class="sxs-lookup"><span data-stu-id="5812c-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="5812c-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="5812c-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="5812c-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="5812c-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5812c-200">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="5812c-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="5812c-201">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5812c-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="5812c-202">型</span><span class="sxs-lookup"><span data-stu-id="5812c-202">Type</span></span>

*   <span data-ttu-id="5812c-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="5812c-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="5812c-204">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-204">Requirements</span></span>

|<span data-ttu-id="5812c-205">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-205">Requirement</span></span>| <span data-ttu-id="5812c-206">値</span><span class="sxs-lookup"><span data-stu-id="5812c-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-207">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-208">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-208">1.0</span></span>|
|[<span data-ttu-id="5812c-209">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-209">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-210">ReadItem</span></span>|
|[<span data-ttu-id="5812c-211">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-211">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-212">読み取り</span><span class="sxs-lookup"><span data-stu-id="5812c-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5812c-213">例</span><span class="sxs-lookup"><span data-stu-id="5812c-213">Example</span></span>

<span data-ttu-id="5812c-214">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="5812c-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
var item = Office.context.mailbox.item;
var outputString = "";

if (item.attachments.length > 0) {
  for (i = 0 ; i < item.attachments.length ; i++) {
    var attachment = item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += attachment.name;
    outputString += "<BR>ID: " + attachment.id;
    outputString += "<BR>contentType: " + attachment.contentType;
    outputString += "<BR>size: " + attachment.size;
    outputString += "<BR>attachmentType: " + attachment.attachmentType;
    outputString += "<BR>isInline: " + attachment.isInline;
  }
}

console.log(outputString);
```

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="5812c-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5812c-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="5812c-216">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="5812c-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="5812c-217">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="5812c-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5812c-218">型</span><span class="sxs-lookup"><span data-stu-id="5812c-218">Type</span></span>

*   [<span data-ttu-id="5812c-219">受信者</span><span class="sxs-lookup"><span data-stu-id="5812c-219">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="5812c-220">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-220">Requirements</span></span>

|<span data-ttu-id="5812c-221">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-221">Requirement</span></span>| <span data-ttu-id="5812c-222">値</span><span class="sxs-lookup"><span data-stu-id="5812c-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-223">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-224">1.1</span><span class="sxs-lookup"><span data-stu-id="5812c-224">1.1</span></span>|
|[<span data-ttu-id="5812c-225">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-225">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-226">ReadItem</span></span>|
|[<span data-ttu-id="5812c-227">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-227">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-228">作成</span><span class="sxs-lookup"><span data-stu-id="5812c-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5812c-229">例</span><span class="sxs-lookup"><span data-stu-id="5812c-229">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="5812c-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="5812c-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="5812c-231">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="5812c-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="5812c-232">型</span><span class="sxs-lookup"><span data-stu-id="5812c-232">Type</span></span>

*   [<span data-ttu-id="5812c-233">Body</span><span class="sxs-lookup"><span data-stu-id="5812c-233">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="5812c-234">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-234">Requirements</span></span>

|<span data-ttu-id="5812c-235">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-235">Requirement</span></span>| <span data-ttu-id="5812c-236">値</span><span class="sxs-lookup"><span data-stu-id="5812c-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-237">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-238">1.1</span><span class="sxs-lookup"><span data-stu-id="5812c-238">1.1</span></span>|
|[<span data-ttu-id="5812c-239">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-239">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-240">ReadItem</span></span>|
|[<span data-ttu-id="5812c-241">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-241">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-242">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5812c-242">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5812c-243">例</span><span class="sxs-lookup"><span data-stu-id="5812c-243">Example</span></span>

<span data-ttu-id="5812c-244">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="5812c-244">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="5812c-245">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="5812c-245">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="5812c-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5812c-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="5812c-247">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="5812c-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="5812c-248">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="5812c-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5812c-249">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="5812c-249">Read mode</span></span>

<span data-ttu-id="5812c-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="5812c-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="5812c-252">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="5812c-252">Compose mode</span></span>

<span data-ttu-id="5812c-253">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="5812c-254">型</span><span class="sxs-lookup"><span data-stu-id="5812c-254">Type</span></span>

*   <span data-ttu-id="5812c-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5812c-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5812c-256">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-256">Requirements</span></span>

|<span data-ttu-id="5812c-257">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-257">Requirement</span></span>| <span data-ttu-id="5812c-258">値</span><span class="sxs-lookup"><span data-stu-id="5812c-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-259">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-260">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-260">1.0</span></span>|
|[<span data-ttu-id="5812c-261">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-261">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-262">ReadItem</span></span>|
|[<span data-ttu-id="5812c-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-263">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-264">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5812c-264">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="5812c-265">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="5812c-265">(nullable) conversationId :String</span></span>

<span data-ttu-id="5812c-266">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="5812c-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="5812c-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="5812c-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="5812c-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="5812c-271">型</span><span class="sxs-lookup"><span data-stu-id="5812c-271">Type</span></span>

*   <span data-ttu-id="5812c-272">String</span><span class="sxs-lookup"><span data-stu-id="5812c-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5812c-273">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-273">Requirements</span></span>

|<span data-ttu-id="5812c-274">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-274">Requirement</span></span>| <span data-ttu-id="5812c-275">値</span><span class="sxs-lookup"><span data-stu-id="5812c-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-276">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-277">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-277">1.0</span></span>|
|[<span data-ttu-id="5812c-278">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-278">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-279">ReadItem</span></span>|
|[<span data-ttu-id="5812c-280">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-280">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-281">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5812c-281">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5812c-282">例</span><span class="sxs-lookup"><span data-stu-id="5812c-282">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="5812c-283">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="5812c-283">dateTimeCreated :Date</span></span>

<span data-ttu-id="5812c-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="5812c-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5812c-286">型</span><span class="sxs-lookup"><span data-stu-id="5812c-286">Type</span></span>

*   <span data-ttu-id="5812c-287">日付</span><span class="sxs-lookup"><span data-stu-id="5812c-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="5812c-288">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-288">Requirements</span></span>

|<span data-ttu-id="5812c-289">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-289">Requirement</span></span>| <span data-ttu-id="5812c-290">値</span><span class="sxs-lookup"><span data-stu-id="5812c-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-291">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-292">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-292">1.0</span></span>|
|[<span data-ttu-id="5812c-293">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-294">ReadItem</span></span>|
|[<span data-ttu-id="5812c-295">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-296">読み取り</span><span class="sxs-lookup"><span data-stu-id="5812c-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5812c-297">例</span><span class="sxs-lookup"><span data-stu-id="5812c-297">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="5812c-298">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="5812c-298">dateTimeModified :Date</span></span>

<span data-ttu-id="5812c-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="5812c-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5812c-301">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="5812c-301">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="5812c-302">型</span><span class="sxs-lookup"><span data-stu-id="5812c-302">Type</span></span>

*   <span data-ttu-id="5812c-303">日付</span><span class="sxs-lookup"><span data-stu-id="5812c-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="5812c-304">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-304">Requirements</span></span>

|<span data-ttu-id="5812c-305">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-305">Requirement</span></span>| <span data-ttu-id="5812c-306">値</span><span class="sxs-lookup"><span data-stu-id="5812c-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-307">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-308">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-308">1.0</span></span>|
|[<span data-ttu-id="5812c-309">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-309">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-310">ReadItem</span></span>|
|[<span data-ttu-id="5812c-311">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-311">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-312">読み取り</span><span class="sxs-lookup"><span data-stu-id="5812c-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5812c-313">例</span><span class="sxs-lookup"><span data-stu-id="5812c-313">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="5812c-314">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="5812c-314">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="5812c-315">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="5812c-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="5812c-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="5812c-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5812c-318">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="5812c-318">Read mode</span></span>

<span data-ttu-id="5812c-319">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-319">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="5812c-320">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="5812c-320">Compose mode</span></span>

<span data-ttu-id="5812c-321">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="5812c-322">[`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5812c-322">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="5812c-323">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="5812c-323">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="5812c-324">型</span><span class="sxs-lookup"><span data-stu-id="5812c-324">Type</span></span>

*   <span data-ttu-id="5812c-325">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="5812c-325">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5812c-326">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-326">Requirements</span></span>

|<span data-ttu-id="5812c-327">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-327">Requirement</span></span>| <span data-ttu-id="5812c-328">値</span><span class="sxs-lookup"><span data-stu-id="5812c-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-329">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-330">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-330">1.0</span></span>|
|[<span data-ttu-id="5812c-331">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-331">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-332">ReadItem</span></span>|
|[<span data-ttu-id="5812c-333">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-333">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-334">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5812c-334">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="5812c-335">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="5812c-335">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="5812c-p112">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="5812c-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="5812c-p113">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="5812c-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="5812c-340">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="5812c-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="5812c-341">型</span><span class="sxs-lookup"><span data-stu-id="5812c-341">Type</span></span>

*   [<span data-ttu-id="5812c-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="5812c-342">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="5812c-343">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-343">Requirements</span></span>

|<span data-ttu-id="5812c-344">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-344">Requirement</span></span>| <span data-ttu-id="5812c-345">値</span><span class="sxs-lookup"><span data-stu-id="5812c-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-346">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-347">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-347">1.0</span></span>|
|[<span data-ttu-id="5812c-348">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-348">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-349">ReadItem</span></span>|
|[<span data-ttu-id="5812c-350">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-350">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-351">読み取り</span><span class="sxs-lookup"><span data-stu-id="5812c-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5812c-352">例</span><span class="sxs-lookup"><span data-stu-id="5812c-352">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="5812c-353">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="5812c-353">internetMessageId :String</span></span>

<span data-ttu-id="5812c-p114">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="5812c-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5812c-356">型</span><span class="sxs-lookup"><span data-stu-id="5812c-356">Type</span></span>

*   <span data-ttu-id="5812c-357">String</span><span class="sxs-lookup"><span data-stu-id="5812c-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5812c-358">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-358">Requirements</span></span>

|<span data-ttu-id="5812c-359">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-359">Requirement</span></span>| <span data-ttu-id="5812c-360">値</span><span class="sxs-lookup"><span data-stu-id="5812c-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-361">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-362">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-362">1.0</span></span>|
|[<span data-ttu-id="5812c-363">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-364">ReadItem</span></span>|
|[<span data-ttu-id="5812c-365">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-366">読み取り</span><span class="sxs-lookup"><span data-stu-id="5812c-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5812c-367">例</span><span class="sxs-lookup"><span data-stu-id="5812c-367">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="5812c-368">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="5812c-368">itemClass :String</span></span>

<span data-ttu-id="5812c-p115">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="5812c-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="5812c-p116">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="5812c-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="5812c-373">型</span><span class="sxs-lookup"><span data-stu-id="5812c-373">Type</span></span> | <span data-ttu-id="5812c-374">説明</span><span class="sxs-lookup"><span data-stu-id="5812c-374">Description</span></span> | <span data-ttu-id="5812c-375">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="5812c-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="5812c-376">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="5812c-376">Appointment items</span></span> | <span data-ttu-id="5812c-377">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="5812c-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="5812c-378">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="5812c-378">Message items</span></span> | <span data-ttu-id="5812c-379">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="5812c-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="5812c-380">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="5812c-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="5812c-381">型</span><span class="sxs-lookup"><span data-stu-id="5812c-381">Type</span></span>

*   <span data-ttu-id="5812c-382">String</span><span class="sxs-lookup"><span data-stu-id="5812c-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5812c-383">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-383">Requirements</span></span>

|<span data-ttu-id="5812c-384">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-384">Requirement</span></span>| <span data-ttu-id="5812c-385">値</span><span class="sxs-lookup"><span data-stu-id="5812c-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-386">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-387">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-387">1.0</span></span>|
|[<span data-ttu-id="5812c-388">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-388">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-389">ReadItem</span></span>|
|[<span data-ttu-id="5812c-390">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-390">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-391">読み取り</span><span class="sxs-lookup"><span data-stu-id="5812c-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5812c-392">例</span><span class="sxs-lookup"><span data-stu-id="5812c-392">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="5812c-393">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="5812c-393">(nullable) itemId :String</span></span>

<span data-ttu-id="5812c-p117">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="5812c-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="5812c-396">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="5812c-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="5812c-397">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="5812c-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="5812c-398">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5812c-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="5812c-399">詳細は、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5812c-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="5812c-p119">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="5812c-402">型</span><span class="sxs-lookup"><span data-stu-id="5812c-402">Type</span></span>

*   <span data-ttu-id="5812c-403">String</span><span class="sxs-lookup"><span data-stu-id="5812c-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5812c-404">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-404">Requirements</span></span>

|<span data-ttu-id="5812c-405">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-405">Requirement</span></span>| <span data-ttu-id="5812c-406">値</span><span class="sxs-lookup"><span data-stu-id="5812c-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-407">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-408">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-408">1.0</span></span>|
|[<span data-ttu-id="5812c-409">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-409">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-410">ReadItem</span></span>|
|[<span data-ttu-id="5812c-411">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-411">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-412">読み取り</span><span class="sxs-lookup"><span data-stu-id="5812c-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5812c-413">例</span><span class="sxs-lookup"><span data-stu-id="5812c-413">Example</span></span>

<span data-ttu-id="5812c-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="5812c-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="5812c-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="5812c-417">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="5812c-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="5812c-418">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="5812c-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="5812c-419">型</span><span class="sxs-lookup"><span data-stu-id="5812c-419">Type</span></span>

*   [<span data-ttu-id="5812c-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="5812c-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="5812c-421">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-421">Requirements</span></span>

|<span data-ttu-id="5812c-422">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-422">Requirement</span></span>| <span data-ttu-id="5812c-423">値</span><span class="sxs-lookup"><span data-stu-id="5812c-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-424">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-425">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-425">1.0</span></span>|
|[<span data-ttu-id="5812c-426">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-427">ReadItem</span></span>|
|[<span data-ttu-id="5812c-428">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-429">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5812c-429">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5812c-430">例</span><span class="sxs-lookup"><span data-stu-id="5812c-430">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="5812c-431">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="5812c-431">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="5812c-432">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="5812c-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5812c-433">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="5812c-433">Read mode</span></span>

<span data-ttu-id="5812c-434">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-434">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="5812c-435">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="5812c-435">Compose mode</span></span>

<span data-ttu-id="5812c-436">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="5812c-437">型</span><span class="sxs-lookup"><span data-stu-id="5812c-437">Type</span></span>

*   <span data-ttu-id="5812c-438">String | [Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="5812c-438">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5812c-439">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-439">Requirements</span></span>

|<span data-ttu-id="5812c-440">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-440">Requirement</span></span>| <span data-ttu-id="5812c-441">値</span><span class="sxs-lookup"><span data-stu-id="5812c-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-442">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-443">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-443">1.0</span></span>|
|[<span data-ttu-id="5812c-444">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-445">ReadItem</span></span>|
|[<span data-ttu-id="5812c-446">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-447">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5812c-447">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="5812c-448">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="5812c-448">normalizedSubject :String</span></span>

<span data-ttu-id="5812c-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="5812c-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="5812c-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="5812c-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="5812c-453">型</span><span class="sxs-lookup"><span data-stu-id="5812c-453">Type</span></span>

*   <span data-ttu-id="5812c-454">String</span><span class="sxs-lookup"><span data-stu-id="5812c-454">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5812c-455">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-455">Requirements</span></span>

|<span data-ttu-id="5812c-456">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-456">Requirement</span></span>| <span data-ttu-id="5812c-457">値</span><span class="sxs-lookup"><span data-stu-id="5812c-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-458">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-459">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-459">1.0</span></span>|
|[<span data-ttu-id="5812c-460">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-460">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-461">ReadItem</span></span>|
|[<span data-ttu-id="5812c-462">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-462">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-463">読み取り</span><span class="sxs-lookup"><span data-stu-id="5812c-463">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5812c-464">例</span><span class="sxs-lookup"><span data-stu-id="5812c-464">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="5812c-465">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="5812c-465">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="5812c-466">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="5812c-466">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="5812c-467">型</span><span class="sxs-lookup"><span data-stu-id="5812c-467">Type</span></span>

*   [<span data-ttu-id="5812c-468">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="5812c-468">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="5812c-469">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-469">Requirements</span></span>

|<span data-ttu-id="5812c-470">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-470">Requirement</span></span>| <span data-ttu-id="5812c-471">値</span><span class="sxs-lookup"><span data-stu-id="5812c-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-472">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-473">1.3</span><span class="sxs-lookup"><span data-stu-id="5812c-473">1.3</span></span>|
|[<span data-ttu-id="5812c-474">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-474">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-475">ReadItem</span></span>|
|[<span data-ttu-id="5812c-476">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-476">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-477">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5812c-477">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5812c-478">例</span><span class="sxs-lookup"><span data-stu-id="5812c-478">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="5812c-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5812c-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="5812c-480">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="5812c-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="5812c-481">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="5812c-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5812c-482">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="5812c-482">Read mode</span></span>

<span data-ttu-id="5812c-483">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="5812c-484">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="5812c-484">Compose mode</span></span>

<span data-ttu-id="5812c-485">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="5812c-486">型</span><span class="sxs-lookup"><span data-stu-id="5812c-486">Type</span></span>

*   <span data-ttu-id="5812c-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5812c-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5812c-488">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-488">Requirements</span></span>

|<span data-ttu-id="5812c-489">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-489">Requirement</span></span>| <span data-ttu-id="5812c-490">値</span><span class="sxs-lookup"><span data-stu-id="5812c-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-491">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-491">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-492">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-492">1.0</span></span>|
|[<span data-ttu-id="5812c-493">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-493">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-494">ReadItem</span></span>|
|[<span data-ttu-id="5812c-495">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-495">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-496">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5812c-496">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="5812c-497">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="5812c-497">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="5812c-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="5812c-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="5812c-500">型</span><span class="sxs-lookup"><span data-stu-id="5812c-500">Type</span></span>

*   [<span data-ttu-id="5812c-501">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="5812c-501">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="5812c-502">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-502">Requirements</span></span>

|<span data-ttu-id="5812c-503">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-503">Requirement</span></span>| <span data-ttu-id="5812c-504">値</span><span class="sxs-lookup"><span data-stu-id="5812c-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-505">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-506">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-506">1.0</span></span>|
|[<span data-ttu-id="5812c-507">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-507">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-508">ReadItem</span></span>|
|[<span data-ttu-id="5812c-509">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-509">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-510">読み取り</span><span class="sxs-lookup"><span data-stu-id="5812c-510">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5812c-511">例</span><span class="sxs-lookup"><span data-stu-id="5812c-511">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="5812c-512">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5812c-512">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="5812c-513">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="5812c-513">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="5812c-514">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="5812c-514">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5812c-515">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="5812c-515">Read mode</span></span>

<span data-ttu-id="5812c-516">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-516">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="5812c-517">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="5812c-517">Compose mode</span></span>

<span data-ttu-id="5812c-518">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-518">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="5812c-519">型</span><span class="sxs-lookup"><span data-stu-id="5812c-519">Type</span></span>

*   <span data-ttu-id="5812c-520">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5812c-520">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5812c-521">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-521">Requirements</span></span>

|<span data-ttu-id="5812c-522">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-522">Requirement</span></span>| <span data-ttu-id="5812c-523">値</span><span class="sxs-lookup"><span data-stu-id="5812c-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-524">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-525">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-525">1.0</span></span>|
|[<span data-ttu-id="5812c-526">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-526">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-527">ReadItem</span></span>|
|[<span data-ttu-id="5812c-528">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-528">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-529">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5812c-529">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="5812c-530">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="5812c-530">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="5812c-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="5812c-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="5812c-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="5812c-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="5812c-535">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="5812c-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="5812c-536">型</span><span class="sxs-lookup"><span data-stu-id="5812c-536">Type</span></span>

*   [<span data-ttu-id="5812c-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="5812c-537">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="5812c-538">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-538">Requirements</span></span>

|<span data-ttu-id="5812c-539">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-539">Requirement</span></span>| <span data-ttu-id="5812c-540">値</span><span class="sxs-lookup"><span data-stu-id="5812c-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-541">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-542">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-542">1.0</span></span>|
|[<span data-ttu-id="5812c-543">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-543">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-544">ReadItem</span></span>|
|[<span data-ttu-id="5812c-545">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-545">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-546">読み取り</span><span class="sxs-lookup"><span data-stu-id="5812c-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5812c-547">例</span><span class="sxs-lookup"><span data-stu-id="5812c-547">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="5812c-548">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="5812c-548">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="5812c-549">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="5812c-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="5812c-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="5812c-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5812c-552">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="5812c-552">Read mode</span></span>

<span data-ttu-id="5812c-553">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-553">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="5812c-554">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="5812c-554">Compose mode</span></span>

<span data-ttu-id="5812c-555">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="5812c-556">[`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5812c-556">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="5812c-557">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="5812c-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="5812c-558">型</span><span class="sxs-lookup"><span data-stu-id="5812c-558">Type</span></span>

*   <span data-ttu-id="5812c-559">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="5812c-559">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5812c-560">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-560">Requirements</span></span>

|<span data-ttu-id="5812c-561">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-561">Requirement</span></span>| <span data-ttu-id="5812c-562">値</span><span class="sxs-lookup"><span data-stu-id="5812c-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-563">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-564">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-564">1.0</span></span>|
|[<span data-ttu-id="5812c-565">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-565">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-566">ReadItem</span></span>|
|[<span data-ttu-id="5812c-567">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-567">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-568">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5812c-568">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="5812c-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="5812c-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="5812c-570">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="5812c-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="5812c-571">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="5812c-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5812c-572">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="5812c-572">Read mode</span></span>

<span data-ttu-id="5812c-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="5812c-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="5812c-575">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="5812c-575">Compose mode</span></span>

<span data-ttu-id="5812c-576">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="5812c-577">型</span><span class="sxs-lookup"><span data-stu-id="5812c-577">Type</span></span>

*   <span data-ttu-id="5812c-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="5812c-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5812c-579">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-579">Requirements</span></span>

|<span data-ttu-id="5812c-580">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-580">Requirement</span></span>| <span data-ttu-id="5812c-581">値</span><span class="sxs-lookup"><span data-stu-id="5812c-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-582">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-583">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-583">1.0</span></span>|
|[<span data-ttu-id="5812c-584">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-584">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-585">ReadItem</span></span>|
|[<span data-ttu-id="5812c-586">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-586">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-587">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5812c-587">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="5812c-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5812c-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="5812c-589">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="5812c-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="5812c-590">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="5812c-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="5812c-591">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="5812c-591">Read mode</span></span>

<span data-ttu-id="5812c-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="5812c-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="5812c-594">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="5812c-594">Compose mode</span></span>

<span data-ttu-id="5812c-595">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="5812c-596">型</span><span class="sxs-lookup"><span data-stu-id="5812c-596">Type</span></span>

*   <span data-ttu-id="5812c-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="5812c-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="5812c-598">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-598">Requirements</span></span>

|<span data-ttu-id="5812c-599">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-599">Requirement</span></span>| <span data-ttu-id="5812c-600">値</span><span class="sxs-lookup"><span data-stu-id="5812c-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-601">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-602">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-602">1.0</span></span>|
|[<span data-ttu-id="5812c-603">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-603">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-604">ReadItem</span></span>|
|[<span data-ttu-id="5812c-605">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-605">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-606">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5812c-606">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="5812c-607">メソッド</span><span class="sxs-lookup"><span data-stu-id="5812c-607">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="5812c-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="5812c-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="5812c-609">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="5812c-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="5812c-610">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="5812c-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="5812c-611">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="5812c-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5812c-612">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5812c-612">Parameters</span></span>

|<span data-ttu-id="5812c-613">名前</span><span class="sxs-lookup"><span data-stu-id="5812c-613">Name</span></span>| <span data-ttu-id="5812c-614">型</span><span class="sxs-lookup"><span data-stu-id="5812c-614">Type</span></span>| <span data-ttu-id="5812c-615">属性</span><span class="sxs-lookup"><span data-stu-id="5812c-615">Attributes</span></span>| <span data-ttu-id="5812c-616">説明</span><span class="sxs-lookup"><span data-stu-id="5812c-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="5812c-617">String</span><span class="sxs-lookup"><span data-stu-id="5812c-617">String</span></span>||<span data-ttu-id="5812c-p132">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="5812c-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="5812c-620">String</span><span class="sxs-lookup"><span data-stu-id="5812c-620">String</span></span>||<span data-ttu-id="5812c-p133">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="5812c-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="5812c-623">Object</span><span class="sxs-lookup"><span data-stu-id="5812c-623">Object</span></span>| <span data-ttu-id="5812c-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-624">&lt;optional&gt;</span></span>|<span data-ttu-id="5812c-625">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="5812c-625">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="5812c-626">Object</span><span class="sxs-lookup"><span data-stu-id="5812c-626">Object</span></span> | <span data-ttu-id="5812c-627">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-627">&lt;optional&gt;</span></span> | <span data-ttu-id="5812c-628">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="5812c-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="5812c-629">Boolean</span><span class="sxs-lookup"><span data-stu-id="5812c-629">Boolean</span></span> | <span data-ttu-id="5812c-630">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-630">&lt;optional&gt;</span></span> | <span data-ttu-id="5812c-631">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="5812c-631">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="5812c-632">function</span><span class="sxs-lookup"><span data-stu-id="5812c-632">function</span></span>| <span data-ttu-id="5812c-633">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-633">&lt;optional&gt;</span></span>|<span data-ttu-id="5812c-634">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="5812c-635">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-635">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="5812c-636">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="5812c-636">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5812c-637">エラー</span><span class="sxs-lookup"><span data-stu-id="5812c-637">Errors</span></span>

| <span data-ttu-id="5812c-638">エラー コード</span><span class="sxs-lookup"><span data-stu-id="5812c-638">Error code</span></span> | <span data-ttu-id="5812c-639">説明</span><span class="sxs-lookup"><span data-stu-id="5812c-639">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="5812c-640">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="5812c-640">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="5812c-641">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="5812c-641">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="5812c-642">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="5812c-642">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5812c-643">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-643">Requirements</span></span>

|<span data-ttu-id="5812c-644">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-644">Requirement</span></span>| <span data-ttu-id="5812c-645">値</span><span class="sxs-lookup"><span data-stu-id="5812c-645">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-646">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-646">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-647">1.1</span><span class="sxs-lookup"><span data-stu-id="5812c-647">1.1</span></span>|
|[<span data-ttu-id="5812c-648">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-648">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-649">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5812c-649">ReadWriteItem</span></span>|
|[<span data-ttu-id="5812c-650">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-650">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-651">作成</span><span class="sxs-lookup"><span data-stu-id="5812c-651">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="5812c-652">例</span><span class="sxs-lookup"><span data-stu-id="5812c-652">Examples</span></span>

```javascript
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="5812c-653">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="5812c-653">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```javascript
Office.context.mailbox.item.addFileAttachmentAsync(
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
        // Do something here.
      });
  });
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="5812c-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="5812c-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="5812c-655">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="5812c-655">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="5812c-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="5812c-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="5812c-659">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="5812c-659">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="5812c-660">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="5812c-660">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5812c-661">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5812c-661">Parameters</span></span>

|<span data-ttu-id="5812c-662">名前</span><span class="sxs-lookup"><span data-stu-id="5812c-662">Name</span></span>| <span data-ttu-id="5812c-663">型</span><span class="sxs-lookup"><span data-stu-id="5812c-663">Type</span></span>| <span data-ttu-id="5812c-664">属性</span><span class="sxs-lookup"><span data-stu-id="5812c-664">Attributes</span></span>| <span data-ttu-id="5812c-665">説明</span><span class="sxs-lookup"><span data-stu-id="5812c-665">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="5812c-666">String</span><span class="sxs-lookup"><span data-stu-id="5812c-666">String</span></span>||<span data-ttu-id="5812c-p135">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="5812c-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="5812c-669">String</span><span class="sxs-lookup"><span data-stu-id="5812c-669">String</span></span>||<span data-ttu-id="5812c-670">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="5812c-670">The subject of the item to be attached.</span></span> <span data-ttu-id="5812c-671">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="5812c-671">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="5812c-672">Object</span><span class="sxs-lookup"><span data-stu-id="5812c-672">Object</span></span>| <span data-ttu-id="5812c-673">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-673">&lt;optional&gt;</span></span>|<span data-ttu-id="5812c-674">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="5812c-674">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5812c-675">Object</span><span class="sxs-lookup"><span data-stu-id="5812c-675">Object</span></span>| <span data-ttu-id="5812c-676">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-676">&lt;optional&gt;</span></span>|<span data-ttu-id="5812c-677">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="5812c-677">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5812c-678">function</span><span class="sxs-lookup"><span data-stu-id="5812c-678">function</span></span>| <span data-ttu-id="5812c-679">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-679">&lt;optional&gt;</span></span>|<span data-ttu-id="5812c-680">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-680">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="5812c-681">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-681">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="5812c-682">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="5812c-682">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5812c-683">エラー</span><span class="sxs-lookup"><span data-stu-id="5812c-683">Errors</span></span>

| <span data-ttu-id="5812c-684">エラー コード</span><span class="sxs-lookup"><span data-stu-id="5812c-684">Error code</span></span> | <span data-ttu-id="5812c-685">説明</span><span class="sxs-lookup"><span data-stu-id="5812c-685">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="5812c-686">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="5812c-686">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5812c-687">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-687">Requirements</span></span>

|<span data-ttu-id="5812c-688">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-688">Requirement</span></span>| <span data-ttu-id="5812c-689">値</span><span class="sxs-lookup"><span data-stu-id="5812c-689">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-690">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-690">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-691">1.1</span><span class="sxs-lookup"><span data-stu-id="5812c-691">1.1</span></span>|
|[<span data-ttu-id="5812c-692">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-692">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-693">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5812c-693">ReadWriteItem</span></span>|
|[<span data-ttu-id="5812c-694">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-694">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-695">作成</span><span class="sxs-lookup"><span data-stu-id="5812c-695">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5812c-696">例</span><span class="sxs-lookup"><span data-stu-id="5812c-696">Example</span></span>

<span data-ttu-id="5812c-697">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-697">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```javascript
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach (shortened for readability).
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="5812c-698">close()</span><span class="sxs-lookup"><span data-stu-id="5812c-698">close()</span></span>

<span data-ttu-id="5812c-699">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="5812c-699">Closes the current item that is being composed.</span></span>

<span data-ttu-id="5812c-p137">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="5812c-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="5812c-702">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-702">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="5812c-703">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="5812c-703">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5812c-704">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-704">Requirements</span></span>

|<span data-ttu-id="5812c-705">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-705">Requirement</span></span>| <span data-ttu-id="5812c-706">値</span><span class="sxs-lookup"><span data-stu-id="5812c-706">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-707">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-707">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-708">1.3</span><span class="sxs-lookup"><span data-stu-id="5812c-708">1.3</span></span>|
|[<span data-ttu-id="5812c-709">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-709">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-710">制限あり</span><span class="sxs-lookup"><span data-stu-id="5812c-710">Restricted</span></span>|
|[<span data-ttu-id="5812c-711">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-711">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-712">新規作成</span><span class="sxs-lookup"><span data-stu-id="5812c-712">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="5812c-713">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="5812c-713">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="5812c-714">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-714">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="5812c-715">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="5812c-715">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5812c-716">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-716">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="5812c-717">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="5812c-717">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="5812c-p138">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="5812c-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5812c-721">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5812c-721">Parameters</span></span>

| <span data-ttu-id="5812c-722">名前</span><span class="sxs-lookup"><span data-stu-id="5812c-722">Name</span></span> | <span data-ttu-id="5812c-723">型</span><span class="sxs-lookup"><span data-stu-id="5812c-723">Type</span></span> | <span data-ttu-id="5812c-724">属性</span><span class="sxs-lookup"><span data-stu-id="5812c-724">Attributes</span></span> | <span data-ttu-id="5812c-725">説明</span><span class="sxs-lookup"><span data-stu-id="5812c-725">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="5812c-726">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="5812c-726">String &#124; Object</span></span>| |<span data-ttu-id="5812c-727">回答フォームの本文を表すテキストと HTML が含まれる文字列。</span><span class="sxs-lookup"><span data-stu-id="5812c-727">A string that contains text and HTML and that represents the body of the reply form.</span></span> <span data-ttu-id="5812c-728">文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="5812c-728">The string is limited to 32 KB.</span></span><br/><span data-ttu-id="5812c-729">**または**</span><span class="sxs-lookup"><span data-stu-id="5812c-729">**OR**</span></span><br/><span data-ttu-id="5812c-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="5812c-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="5812c-732">String</span><span class="sxs-lookup"><span data-stu-id="5812c-732">String</span></span> | <span data-ttu-id="5812c-733">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-733">&lt;optional&gt;</span></span> | <span data-ttu-id="5812c-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="5812c-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="5812c-736">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-736">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="5812c-737">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-737">&lt;optional&gt;</span></span> | <span data-ttu-id="5812c-738">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="5812c-738">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="5812c-739">String</span><span class="sxs-lookup"><span data-stu-id="5812c-739">String</span></span> | | <span data-ttu-id="5812c-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="5812c-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="5812c-742">String</span><span class="sxs-lookup"><span data-stu-id="5812c-742">String</span></span> | | <span data-ttu-id="5812c-743">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="5812c-743">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="5812c-744">String</span><span class="sxs-lookup"><span data-stu-id="5812c-744">String</span></span> | | <span data-ttu-id="5812c-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="5812c-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="5812c-747">ブール値</span><span class="sxs-lookup"><span data-stu-id="5812c-747">Boolean</span></span> | | <span data-ttu-id="5812c-p144">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="5812c-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="5812c-750">String</span><span class="sxs-lookup"><span data-stu-id="5812c-750">String</span></span> | | <span data-ttu-id="5812c-p145">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="5812c-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="5812c-754">function</span><span class="sxs-lookup"><span data-stu-id="5812c-754">function</span></span> | <span data-ttu-id="5812c-755">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-755">&lt;optional&gt;</span></span> | <span data-ttu-id="5812c-756">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5812c-757">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-757">Requirements</span></span>

|<span data-ttu-id="5812c-758">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-758">Requirement</span></span>| <span data-ttu-id="5812c-759">値</span><span class="sxs-lookup"><span data-stu-id="5812c-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-760">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-760">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-761">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-761">1.0</span></span>|
|[<span data-ttu-id="5812c-762">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-762">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-763">ReadItem</span></span>|
|[<span data-ttu-id="5812c-764">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-764">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-765">読み取り</span><span class="sxs-lookup"><span data-stu-id="5812c-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="5812c-766">例</span><span class="sxs-lookup"><span data-stu-id="5812c-766">Examples</span></span>

<span data-ttu-id="5812c-767">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="5812c-767">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="5812c-768">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="5812c-768">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="5812c-769">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="5812c-769">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="5812c-770">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="5812c-770">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="5812c-771">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="5812c-771">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="5812c-772">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="5812c-772">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="5812c-773">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="5812c-773">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="5812c-774">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-774">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="5812c-775">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="5812c-775">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5812c-776">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-776">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="5812c-777">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="5812c-777">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="5812c-p146">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="5812c-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5812c-781">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5812c-781">Parameters</span></span>

| <span data-ttu-id="5812c-782">名前</span><span class="sxs-lookup"><span data-stu-id="5812c-782">Name</span></span> | <span data-ttu-id="5812c-783">型</span><span class="sxs-lookup"><span data-stu-id="5812c-783">Type</span></span> | <span data-ttu-id="5812c-784">属性</span><span class="sxs-lookup"><span data-stu-id="5812c-784">Attributes</span></span> | <span data-ttu-id="5812c-785">説明</span><span class="sxs-lookup"><span data-stu-id="5812c-785">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="5812c-786">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="5812c-786">String &#124; Object</span></span>| | <span data-ttu-id="5812c-787">回答フォームの本文を表すテキストと HTML が含まれる文字列。</span><span class="sxs-lookup"><span data-stu-id="5812c-787">A string that contains text and HTML and that represents the body of the reply form.</span></span> <span data-ttu-id="5812c-788">文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="5812c-788">The string is limited to 32 KB.</span></span><br/><span data-ttu-id="5812c-789">**または**</span><span class="sxs-lookup"><span data-stu-id="5812c-789">**OR**</span></span><br/><span data-ttu-id="5812c-p148">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="5812c-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="5812c-792">String</span><span class="sxs-lookup"><span data-stu-id="5812c-792">String</span></span> | <span data-ttu-id="5812c-793">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-793">&lt;optional&gt;</span></span> | <span data-ttu-id="5812c-p149">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="5812c-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="5812c-796">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-796">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="5812c-797">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-797">&lt;optional&gt;</span></span> | <span data-ttu-id="5812c-798">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="5812c-798">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="5812c-799">String</span><span class="sxs-lookup"><span data-stu-id="5812c-799">String</span></span> | | <span data-ttu-id="5812c-p150">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="5812c-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="5812c-802">String</span><span class="sxs-lookup"><span data-stu-id="5812c-802">String</span></span> | | <span data-ttu-id="5812c-803">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="5812c-803">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="5812c-804">String</span><span class="sxs-lookup"><span data-stu-id="5812c-804">String</span></span> | | <span data-ttu-id="5812c-p151">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="5812c-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="5812c-807">ブール値</span><span class="sxs-lookup"><span data-stu-id="5812c-807">Boolean</span></span> | | <span data-ttu-id="5812c-p152">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="5812c-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="5812c-810">String</span><span class="sxs-lookup"><span data-stu-id="5812c-810">String</span></span> | | <span data-ttu-id="5812c-p153">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="5812c-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="5812c-814">function</span><span class="sxs-lookup"><span data-stu-id="5812c-814">function</span></span> | <span data-ttu-id="5812c-815">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-815">&lt;optional&gt;</span></span> | <span data-ttu-id="5812c-816">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-816">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5812c-817">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-817">Requirements</span></span>

|<span data-ttu-id="5812c-818">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-818">Requirement</span></span>| <span data-ttu-id="5812c-819">値</span><span class="sxs-lookup"><span data-stu-id="5812c-819">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-820">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-820">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-821">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-821">1.0</span></span>|
|[<span data-ttu-id="5812c-822">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-822">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-823">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-823">ReadItem</span></span>|
|[<span data-ttu-id="5812c-824">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-824">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-825">読み取り</span><span class="sxs-lookup"><span data-stu-id="5812c-825">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="5812c-826">例</span><span class="sxs-lookup"><span data-stu-id="5812c-826">Examples</span></span>

<span data-ttu-id="5812c-827">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="5812c-827">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="5812c-828">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="5812c-828">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="5812c-829">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="5812c-829">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="5812c-830">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="5812c-830">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="5812c-831">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="5812c-831">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="5812c-832">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="5812c-832">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="5812c-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="5812c-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="5812c-834">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="5812c-834">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="5812c-835">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="5812c-835">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5812c-836">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-836">Requirements</span></span>

|<span data-ttu-id="5812c-837">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-837">Requirement</span></span>| <span data-ttu-id="5812c-838">値</span><span class="sxs-lookup"><span data-stu-id="5812c-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-839">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-840">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-840">1.0</span></span>|
|[<span data-ttu-id="5812c-841">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-841">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-842">ReadItem</span></span>|
|[<span data-ttu-id="5812c-843">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-843">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-844">読み取り</span><span class="sxs-lookup"><span data-stu-id="5812c-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5812c-845">戻り値:</span><span class="sxs-lookup"><span data-stu-id="5812c-845">Returns:</span></span>

<span data-ttu-id="5812c-846">型:[Entities](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="5812c-846">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="5812c-847">例</span><span class="sxs-lookup"><span data-stu-id="5812c-847">Example</span></span>

<span data-ttu-id="5812c-848">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="5812c-848">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="5812c-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="5812c-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="5812c-850">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="5812c-850">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="5812c-851">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="5812c-851">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5812c-852">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5812c-852">Parameters</span></span>

|<span data-ttu-id="5812c-853">名前</span><span class="sxs-lookup"><span data-stu-id="5812c-853">Name</span></span>| <span data-ttu-id="5812c-854">型</span><span class="sxs-lookup"><span data-stu-id="5812c-854">Type</span></span>| <span data-ttu-id="5812c-855">説明</span><span class="sxs-lookup"><span data-stu-id="5812c-855">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="5812c-856">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="5812c-856">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="5812c-857">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="5812c-857">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5812c-858">Requirements</span><span class="sxs-lookup"><span data-stu-id="5812c-858">Requirements</span></span>

|<span data-ttu-id="5812c-859">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-859">Requirement</span></span>| <span data-ttu-id="5812c-860">値</span><span class="sxs-lookup"><span data-stu-id="5812c-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-861">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-862">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-862">1.0</span></span>|
|[<span data-ttu-id="5812c-863">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-863">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-864">制限あり</span><span class="sxs-lookup"><span data-stu-id="5812c-864">Restricted</span></span>|
|[<span data-ttu-id="5812c-865">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-865">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-866">読み取り</span><span class="sxs-lookup"><span data-stu-id="5812c-866">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5812c-867">戻り値:</span><span class="sxs-lookup"><span data-stu-id="5812c-867">Returns:</span></span>

<span data-ttu-id="5812c-868">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-868">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="5812c-869">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-869">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="5812c-870">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="5812c-870">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="5812c-871">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="5812c-871">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="5812c-872">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="5812c-872">Value of `entityType`</span></span> | <span data-ttu-id="5812c-873">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="5812c-873">Type of objects in returned array</span></span> | <span data-ttu-id="5812c-874">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="5812c-874">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="5812c-875">文字列</span><span class="sxs-lookup"><span data-stu-id="5812c-875">String</span></span> | <span data-ttu-id="5812c-876">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="5812c-876">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="5812c-877">連絡先</span><span class="sxs-lookup"><span data-stu-id="5812c-877">Contact</span></span> | <span data-ttu-id="5812c-878">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5812c-878">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="5812c-879">文字列</span><span class="sxs-lookup"><span data-stu-id="5812c-879">String</span></span> | <span data-ttu-id="5812c-880">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5812c-880">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="5812c-881">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="5812c-881">MeetingSuggestion</span></span> | <span data-ttu-id="5812c-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5812c-882">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="5812c-883">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="5812c-883">PhoneNumber</span></span> | <span data-ttu-id="5812c-884">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="5812c-884">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="5812c-885">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="5812c-885">TaskSuggestion</span></span> | <span data-ttu-id="5812c-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="5812c-886">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="5812c-887">文字列</span><span class="sxs-lookup"><span data-stu-id="5812c-887">String</span></span> | <span data-ttu-id="5812c-888">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="5812c-888">**Restricted**</span></span> |

<span data-ttu-id="5812c-889">型:Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="5812c-889">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="5812c-890">例</span><span class="sxs-lookup"><span data-stu-id="5812c-890">Example</span></span>

<span data-ttu-id="5812c-891">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="5812c-891">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="5812c-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="5812c-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="5812c-893">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-893">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="5812c-894">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="5812c-894">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5812c-895">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-895">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5812c-896">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5812c-896">Parameters</span></span>

|<span data-ttu-id="5812c-897">名前</span><span class="sxs-lookup"><span data-stu-id="5812c-897">Name</span></span>| <span data-ttu-id="5812c-898">型</span><span class="sxs-lookup"><span data-stu-id="5812c-898">Type</span></span>| <span data-ttu-id="5812c-899">説明</span><span class="sxs-lookup"><span data-stu-id="5812c-899">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="5812c-900">String</span><span class="sxs-lookup"><span data-stu-id="5812c-900">String</span></span>|<span data-ttu-id="5812c-901">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="5812c-901">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5812c-902">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-902">Requirements</span></span>

|<span data-ttu-id="5812c-903">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-903">Requirement</span></span>| <span data-ttu-id="5812c-904">値</span><span class="sxs-lookup"><span data-stu-id="5812c-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-905">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-905">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-906">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-906">1.0</span></span>|
|[<span data-ttu-id="5812c-907">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-908">ReadItem</span></span>|
|[<span data-ttu-id="5812c-909">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-910">読み取り</span><span class="sxs-lookup"><span data-stu-id="5812c-910">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5812c-911">戻り値:</span><span class="sxs-lookup"><span data-stu-id="5812c-911">Returns:</span></span>

<span data-ttu-id="5812c-p155">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="5812c-914">型:Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="5812c-914">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="5812c-915">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="5812c-915">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="5812c-916">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-916">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="5812c-917">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="5812c-917">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5812c-p156">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="5812c-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="5812c-921">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="5812c-921">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="5812c-922">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="5812c-922">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="5812c-p157">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="5812c-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="5812c-926">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-926">Requirements</span></span>

|<span data-ttu-id="5812c-927">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-927">Requirement</span></span>| <span data-ttu-id="5812c-928">値</span><span class="sxs-lookup"><span data-stu-id="5812c-928">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-929">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-929">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-930">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-930">1.0</span></span>|
|[<span data-ttu-id="5812c-931">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-931">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-932">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-932">ReadItem</span></span>|
|[<span data-ttu-id="5812c-933">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-933">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-934">読み取り</span><span class="sxs-lookup"><span data-stu-id="5812c-934">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5812c-935">戻り値:</span><span class="sxs-lookup"><span data-stu-id="5812c-935">Returns:</span></span>

<span data-ttu-id="5812c-p158">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="5812c-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="5812c-938">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="5812c-938">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="5812c-939">Object</span><span class="sxs-lookup"><span data-stu-id="5812c-939">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="5812c-940">例</span><span class="sxs-lookup"><span data-stu-id="5812c-940">Example</span></span>

<span data-ttu-id="5812c-941">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="5812c-941">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="5812c-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="5812c-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="5812c-943">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-943">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="5812c-944">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="5812c-944">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5812c-945">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="5812c-945">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="5812c-p159">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="5812c-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5812c-948">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5812c-948">Parameters</span></span>

|<span data-ttu-id="5812c-949">名前</span><span class="sxs-lookup"><span data-stu-id="5812c-949">Name</span></span>| <span data-ttu-id="5812c-950">型</span><span class="sxs-lookup"><span data-stu-id="5812c-950">Type</span></span>| <span data-ttu-id="5812c-951">説明</span><span class="sxs-lookup"><span data-stu-id="5812c-951">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="5812c-952">String</span><span class="sxs-lookup"><span data-stu-id="5812c-952">String</span></span>|<span data-ttu-id="5812c-953">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="5812c-953">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5812c-954">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-954">Requirements</span></span>

|<span data-ttu-id="5812c-955">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-955">Requirement</span></span>| <span data-ttu-id="5812c-956">値</span><span class="sxs-lookup"><span data-stu-id="5812c-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-957">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-957">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-958">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-958">1.0</span></span>|
|[<span data-ttu-id="5812c-959">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-960">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-960">ReadItem</span></span>|
|[<span data-ttu-id="5812c-961">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-962">読み取り</span><span class="sxs-lookup"><span data-stu-id="5812c-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="5812c-963">戻り値:</span><span class="sxs-lookup"><span data-stu-id="5812c-963">Returns:</span></span>

<span data-ttu-id="5812c-964">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="5812c-964">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="5812c-965">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="5812c-965">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="5812c-966">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="5812c-966">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="5812c-967">例</span><span class="sxs-lookup"><span data-stu-id="5812c-967">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="5812c-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="5812c-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="5812c-969">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-969">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="5812c-p160">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5812c-972">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5812c-972">Parameters</span></span>

|<span data-ttu-id="5812c-973">名前</span><span class="sxs-lookup"><span data-stu-id="5812c-973">Name</span></span>| <span data-ttu-id="5812c-974">型</span><span class="sxs-lookup"><span data-stu-id="5812c-974">Type</span></span>| <span data-ttu-id="5812c-975">属性</span><span class="sxs-lookup"><span data-stu-id="5812c-975">Attributes</span></span>| <span data-ttu-id="5812c-976">説明</span><span class="sxs-lookup"><span data-stu-id="5812c-976">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="5812c-977">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="5812c-977">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="5812c-p161">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="5812c-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="5812c-981">Object</span><span class="sxs-lookup"><span data-stu-id="5812c-981">Object</span></span>| <span data-ttu-id="5812c-982">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-982">&lt;optional&gt;</span></span>|<span data-ttu-id="5812c-983">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="5812c-983">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5812c-984">Object</span><span class="sxs-lookup"><span data-stu-id="5812c-984">Object</span></span>| <span data-ttu-id="5812c-985">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-985">&lt;optional&gt;</span></span>|<span data-ttu-id="5812c-986">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="5812c-986">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5812c-987">function</span><span class="sxs-lookup"><span data-stu-id="5812c-987">function</span></span>||<span data-ttu-id="5812c-988">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-988">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5812c-989">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="5812c-989">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="5812c-990">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="5812c-990">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5812c-991">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-991">Requirements</span></span>

|<span data-ttu-id="5812c-992">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-992">Requirement</span></span>| <span data-ttu-id="5812c-993">値</span><span class="sxs-lookup"><span data-stu-id="5812c-993">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-994">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-994">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-995">1.2</span><span class="sxs-lookup"><span data-stu-id="5812c-995">1.2</span></span>|
|[<span data-ttu-id="5812c-996">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-996">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-997">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5812c-997">ReadWriteItem</span></span>|
|[<span data-ttu-id="5812c-998">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-998">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-999">作成</span><span class="sxs-lookup"><span data-stu-id="5812c-999">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="5812c-1000">戻り値:</span><span class="sxs-lookup"><span data-stu-id="5812c-1000">Returns:</span></span>

<span data-ttu-id="5812c-1001">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="5812c-1001">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="5812c-1002">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="5812c-1002">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="5812c-1003">String</span><span class="sxs-lookup"><span data-stu-id="5812c-1003">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="5812c-1004">例</span><span class="sxs-lookup"><span data-stu-id="5812c-1004">Example</span></span>

```javascript
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
  // Check for errors.
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="5812c-1005">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="5812c-1005">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="5812c-1006">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="5812c-1006">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="5812c-p163">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="5812c-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5812c-1010">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5812c-1010">Parameters</span></span>

|<span data-ttu-id="5812c-1011">名前</span><span class="sxs-lookup"><span data-stu-id="5812c-1011">Name</span></span>| <span data-ttu-id="5812c-1012">型</span><span class="sxs-lookup"><span data-stu-id="5812c-1012">Type</span></span>| <span data-ttu-id="5812c-1013">属性</span><span class="sxs-lookup"><span data-stu-id="5812c-1013">Attributes</span></span>| <span data-ttu-id="5812c-1014">説明</span><span class="sxs-lookup"><span data-stu-id="5812c-1014">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="5812c-1015">function</span><span class="sxs-lookup"><span data-stu-id="5812c-1015">function</span></span>||<span data-ttu-id="5812c-1016">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5812c-1017">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-1017">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="5812c-1018">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="5812c-1018">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="5812c-1019">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="5812c-1019">Object</span></span>| <span data-ttu-id="5812c-1020">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-1020">&lt;optional&gt;</span></span>|<span data-ttu-id="5812c-1021">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="5812c-1021">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="5812c-1022">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="5812c-1022">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5812c-1023">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-1023">Requirements</span></span>

|<span data-ttu-id="5812c-1024">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-1024">Requirement</span></span>| <span data-ttu-id="5812c-1025">値</span><span class="sxs-lookup"><span data-stu-id="5812c-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-1026">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-1026">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="5812c-1027">1.0</span></span>|
|[<span data-ttu-id="5812c-1028">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-1028">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5812c-1029">ReadItem</span></span>|
|[<span data-ttu-id="5812c-1030">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-1030">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-1031">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="5812c-1031">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5812c-1032">例</span><span class="sxs-lookup"><span data-stu-id="5812c-1032">Example</span></span>

<span data-ttu-id="5812c-p166">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="5812c-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(customPropsCallback);
  });
};

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="5812c-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="5812c-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="5812c-1037">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="5812c-1037">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="5812c-p167">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="5812c-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5812c-1042">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5812c-1042">Parameters</span></span>

|<span data-ttu-id="5812c-1043">名前</span><span class="sxs-lookup"><span data-stu-id="5812c-1043">Name</span></span>| <span data-ttu-id="5812c-1044">型</span><span class="sxs-lookup"><span data-stu-id="5812c-1044">Type</span></span>| <span data-ttu-id="5812c-1045">属性</span><span class="sxs-lookup"><span data-stu-id="5812c-1045">Attributes</span></span>| <span data-ttu-id="5812c-1046">説明</span><span class="sxs-lookup"><span data-stu-id="5812c-1046">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="5812c-1047">String</span><span class="sxs-lookup"><span data-stu-id="5812c-1047">String</span></span>||<span data-ttu-id="5812c-1048">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="5812c-1048">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="5812c-1049">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="5812c-1049">Object</span></span>| <span data-ttu-id="5812c-1050">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-1050">&lt;optional&gt;</span></span>|<span data-ttu-id="5812c-1051">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="5812c-1051">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5812c-1052">Object</span><span class="sxs-lookup"><span data-stu-id="5812c-1052">Object</span></span>| <span data-ttu-id="5812c-1053">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-1053">&lt;optional&gt;</span></span>|<span data-ttu-id="5812c-1054">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="5812c-1054">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5812c-1055">function</span><span class="sxs-lookup"><span data-stu-id="5812c-1055">function</span></span>| <span data-ttu-id="5812c-1056">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-1056">&lt;optional&gt;</span></span>|<span data-ttu-id="5812c-1057">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-1057">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="5812c-1058">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="5812c-1058">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="5812c-1059">エラー</span><span class="sxs-lookup"><span data-stu-id="5812c-1059">Errors</span></span>

| <span data-ttu-id="5812c-1060">エラー コード</span><span class="sxs-lookup"><span data-stu-id="5812c-1060">Error code</span></span> | <span data-ttu-id="5812c-1061">説明</span><span class="sxs-lookup"><span data-stu-id="5812c-1061">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="5812c-1062">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="5812c-1062">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5812c-1063">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-1063">Requirements</span></span>

|<span data-ttu-id="5812c-1064">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-1064">Requirement</span></span>| <span data-ttu-id="5812c-1065">値</span><span class="sxs-lookup"><span data-stu-id="5812c-1065">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-1066">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-1066">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-1067">1.1</span><span class="sxs-lookup"><span data-stu-id="5812c-1067">1.1</span></span>|
|[<span data-ttu-id="5812c-1068">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-1068">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-1069">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5812c-1069">ReadWriteItem</span></span>|
|[<span data-ttu-id="5812c-1070">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-1070">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-1071">作成</span><span class="sxs-lookup"><span data-stu-id="5812c-1071">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5812c-1072">例</span><span class="sxs-lookup"><span data-stu-id="5812c-1072">Example</span></span>

<span data-ttu-id="5812c-1073">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="5812c-1073">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="5812c-1074">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="5812c-1074">saveAsync([options], callback)</span></span>

<span data-ttu-id="5812c-1075">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="5812c-1075">Asynchronously saves an item.</span></span>

<span data-ttu-id="5812c-p168">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-p168">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="5812c-1079">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="5812c-1079">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="5812c-1080">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-1080">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="5812c-p170">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-p170">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="5812c-1084">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="5812c-1084">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="5812c-1085">Mac Outlook では、新規作成モードの会議で `saveAsync` をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="5812c-1085">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="5812c-1086">Mac Outlook では、会議で `saveAsync` を呼び出すとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-1086">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="5812c-1087">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-1087">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5812c-1088">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5812c-1088">Parameters</span></span>

|<span data-ttu-id="5812c-1089">名前</span><span class="sxs-lookup"><span data-stu-id="5812c-1089">Name</span></span>| <span data-ttu-id="5812c-1090">型</span><span class="sxs-lookup"><span data-stu-id="5812c-1090">Type</span></span>| <span data-ttu-id="5812c-1091">属性</span><span class="sxs-lookup"><span data-stu-id="5812c-1091">Attributes</span></span>| <span data-ttu-id="5812c-1092">説明</span><span class="sxs-lookup"><span data-stu-id="5812c-1092">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="5812c-1093">Object</span><span class="sxs-lookup"><span data-stu-id="5812c-1093">Object</span></span>| <span data-ttu-id="5812c-1094">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-1094">&lt;optional&gt;</span></span>|<span data-ttu-id="5812c-1095">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="5812c-1095">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5812c-1096">Object</span><span class="sxs-lookup"><span data-stu-id="5812c-1096">Object</span></span>| <span data-ttu-id="5812c-1097">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="5812c-1098">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="5812c-1098">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="5812c-1099">function</span><span class="sxs-lookup"><span data-stu-id="5812c-1099">function</span></span>||<span data-ttu-id="5812c-1100">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-1100">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="5812c-1101">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-1101">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5812c-1102">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-1102">Requirements</span></span>

|<span data-ttu-id="5812c-1103">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-1103">Requirement</span></span>| <span data-ttu-id="5812c-1104">値</span><span class="sxs-lookup"><span data-stu-id="5812c-1104">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-1105">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-1105">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-1106">1.3</span><span class="sxs-lookup"><span data-stu-id="5812c-1106">1.3</span></span>|
|[<span data-ttu-id="5812c-1107">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-1107">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-1108">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5812c-1108">ReadWriteItem</span></span>|
|[<span data-ttu-id="5812c-1109">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-1109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-1110">作成</span><span class="sxs-lookup"><span data-stu-id="5812c-1110">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="5812c-1111">例</span><span class="sxs-lookup"><span data-stu-id="5812c-1111">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="5812c-p172">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="5812c-p172">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="5812c-1114">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="5812c-1114">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="5812c-1115">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="5812c-1115">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="5812c-p173">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="5812c-p173">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="5812c-1119">パラメーター</span><span class="sxs-lookup"><span data-stu-id="5812c-1119">Parameters</span></span>

|<span data-ttu-id="5812c-1120">名前</span><span class="sxs-lookup"><span data-stu-id="5812c-1120">Name</span></span>| <span data-ttu-id="5812c-1121">型</span><span class="sxs-lookup"><span data-stu-id="5812c-1121">Type</span></span>| <span data-ttu-id="5812c-1122">属性</span><span class="sxs-lookup"><span data-stu-id="5812c-1122">Attributes</span></span>| <span data-ttu-id="5812c-1123">説明</span><span class="sxs-lookup"><span data-stu-id="5812c-1123">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="5812c-1124">String</span><span class="sxs-lookup"><span data-stu-id="5812c-1124">String</span></span>||<span data-ttu-id="5812c-p174">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="5812c-p174">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="5812c-1128">Object</span><span class="sxs-lookup"><span data-stu-id="5812c-1128">Object</span></span>| <span data-ttu-id="5812c-1129">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-1129">&lt;optional&gt;</span></span>|<span data-ttu-id="5812c-1130">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="5812c-1130">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="5812c-1131">Object</span><span class="sxs-lookup"><span data-stu-id="5812c-1131">Object</span></span>| <span data-ttu-id="5812c-1132">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-1132">&lt;optional&gt;</span></span>|<span data-ttu-id="5812c-1133">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="5812c-1133">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="5812c-1134">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="5812c-1134">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="5812c-1135">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="5812c-1135">&lt;optional&gt;</span></span>|<span data-ttu-id="5812c-p175">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-p175">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="5812c-p176">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-p176">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="5812c-1140">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-1140">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="5812c-1141">function</span><span class="sxs-lookup"><span data-stu-id="5812c-1141">function</span></span>||<span data-ttu-id="5812c-1142">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="5812c-1142">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5812c-1143">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-1143">Requirements</span></span>

|<span data-ttu-id="5812c-1144">要件</span><span class="sxs-lookup"><span data-stu-id="5812c-1144">Requirement</span></span>| <span data-ttu-id="5812c-1145">値</span><span class="sxs-lookup"><span data-stu-id="5812c-1145">Value</span></span>|
|---|---|
|[<span data-ttu-id="5812c-1146">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="5812c-1146">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5812c-1147">1.2</span><span class="sxs-lookup"><span data-stu-id="5812c-1147">1.2</span></span>|
|[<span data-ttu-id="5812c-1148">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="5812c-1148">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5812c-1149">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="5812c-1149">ReadWriteItem</span></span>|
|[<span data-ttu-id="5812c-1150">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="5812c-1150">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5812c-1151">作成</span><span class="sxs-lookup"><span data-stu-id="5812c-1151">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="5812c-1152">例</span><span class="sxs-lookup"><span data-stu-id="5812c-1152">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
