---
title: Office. メールボックス-要件セット1.2
description: ''
ms.date: 10/23/2019
localization_priority: Normal
ms.openlocfilehash: e83b80ee2d71913d959ddfaedf5bb80c208d83ef
ms.sourcegitcommit: 5ba325cc88183a3f230cd89d615fd49c695addcf
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/24/2019
ms.locfileid: "37682565"
---
# <a name="item"></a><span data-ttu-id="c9eda-102">item</span><span class="sxs-lookup"><span data-stu-id="c9eda-102">item</span></span>

### <span data-ttu-id="c9eda-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="c9eda-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="c9eda-p102">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9eda-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9eda-107">Requirements</span></span>

|<span data-ttu-id="c9eda-108">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-108">Requirement</span></span>| <span data-ttu-id="c9eda-109">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-111">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-111">1.0</span></span>|
|[<span data-ttu-id="c9eda-112">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-113">制限あり</span><span class="sxs-lookup"><span data-stu-id="c9eda-113">Restricted</span></span>|
|[<span data-ttu-id="c9eda-114">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-115">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c9eda-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c9eda-116">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="c9eda-116">Members and methods</span></span>

| <span data-ttu-id="c9eda-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="c9eda-117">Member</span></span> | <span data-ttu-id="c9eda-118">種類</span><span class="sxs-lookup"><span data-stu-id="c9eda-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c9eda-119">attachments</span><span class="sxs-lookup"><span data-stu-id="c9eda-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="c9eda-120">Member</span><span class="sxs-lookup"><span data-stu-id="c9eda-120">Member</span></span> |
| [<span data-ttu-id="c9eda-121">bcc</span><span class="sxs-lookup"><span data-stu-id="c9eda-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="c9eda-122">Member</span><span class="sxs-lookup"><span data-stu-id="c9eda-122">Member</span></span> |
| [<span data-ttu-id="c9eda-123">body</span><span class="sxs-lookup"><span data-stu-id="c9eda-123">body</span></span>](#body-body) | <span data-ttu-id="c9eda-124">Member</span><span class="sxs-lookup"><span data-stu-id="c9eda-124">Member</span></span> |
| [<span data-ttu-id="c9eda-125">cc</span><span class="sxs-lookup"><span data-stu-id="c9eda-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="c9eda-126">Member</span><span class="sxs-lookup"><span data-stu-id="c9eda-126">Member</span></span> |
| [<span data-ttu-id="c9eda-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="c9eda-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="c9eda-128">Member</span><span class="sxs-lookup"><span data-stu-id="c9eda-128">Member</span></span> |
| [<span data-ttu-id="c9eda-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="c9eda-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="c9eda-130">Member</span><span class="sxs-lookup"><span data-stu-id="c9eda-130">Member</span></span> |
| [<span data-ttu-id="c9eda-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="c9eda-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="c9eda-132">Member</span><span class="sxs-lookup"><span data-stu-id="c9eda-132">Member</span></span> |
| [<span data-ttu-id="c9eda-133">end</span><span class="sxs-lookup"><span data-stu-id="c9eda-133">end</span></span>](#end-datetime) | <span data-ttu-id="c9eda-134">Member</span><span class="sxs-lookup"><span data-stu-id="c9eda-134">Member</span></span> |
| [<span data-ttu-id="c9eda-135">from</span><span class="sxs-lookup"><span data-stu-id="c9eda-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="c9eda-136">Member</span><span class="sxs-lookup"><span data-stu-id="c9eda-136">Member</span></span> |
| [<span data-ttu-id="c9eda-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="c9eda-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="c9eda-138">Member</span><span class="sxs-lookup"><span data-stu-id="c9eda-138">Member</span></span> |
| [<span data-ttu-id="c9eda-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="c9eda-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="c9eda-140">Member</span><span class="sxs-lookup"><span data-stu-id="c9eda-140">Member</span></span> |
| [<span data-ttu-id="c9eda-141">itemId</span><span class="sxs-lookup"><span data-stu-id="c9eda-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="c9eda-142">Member</span><span class="sxs-lookup"><span data-stu-id="c9eda-142">Member</span></span> |
| [<span data-ttu-id="c9eda-143">itemType</span><span class="sxs-lookup"><span data-stu-id="c9eda-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="c9eda-144">Member</span><span class="sxs-lookup"><span data-stu-id="c9eda-144">Member</span></span> |
| [<span data-ttu-id="c9eda-145">location</span><span class="sxs-lookup"><span data-stu-id="c9eda-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="c9eda-146">Member</span><span class="sxs-lookup"><span data-stu-id="c9eda-146">Member</span></span> |
| [<span data-ttu-id="c9eda-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="c9eda-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="c9eda-148">Member</span><span class="sxs-lookup"><span data-stu-id="c9eda-148">Member</span></span> |
| [<span data-ttu-id="c9eda-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="c9eda-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="c9eda-150">Member</span><span class="sxs-lookup"><span data-stu-id="c9eda-150">Member</span></span> |
| [<span data-ttu-id="c9eda-151">organizer</span><span class="sxs-lookup"><span data-stu-id="c9eda-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="c9eda-152">Member</span><span class="sxs-lookup"><span data-stu-id="c9eda-152">Member</span></span> |
| [<span data-ttu-id="c9eda-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="c9eda-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="c9eda-154">Member</span><span class="sxs-lookup"><span data-stu-id="c9eda-154">Member</span></span> |
| [<span data-ttu-id="c9eda-155">sender</span><span class="sxs-lookup"><span data-stu-id="c9eda-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="c9eda-156">Member</span><span class="sxs-lookup"><span data-stu-id="c9eda-156">Member</span></span> |
| [<span data-ttu-id="c9eda-157">start</span><span class="sxs-lookup"><span data-stu-id="c9eda-157">start</span></span>](#start-datetime) | <span data-ttu-id="c9eda-158">Member</span><span class="sxs-lookup"><span data-stu-id="c9eda-158">Member</span></span> |
| [<span data-ttu-id="c9eda-159">subject</span><span class="sxs-lookup"><span data-stu-id="c9eda-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="c9eda-160">メンバー</span><span class="sxs-lookup"><span data-stu-id="c9eda-160">Member</span></span> |
| [<span data-ttu-id="c9eda-161">to</span><span class="sxs-lookup"><span data-stu-id="c9eda-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="c9eda-162">メンバー</span><span class="sxs-lookup"><span data-stu-id="c9eda-162">Member</span></span> |
| [<span data-ttu-id="c9eda-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c9eda-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="c9eda-164">Method</span><span class="sxs-lookup"><span data-stu-id="c9eda-164">Method</span></span> |
| [<span data-ttu-id="c9eda-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c9eda-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="c9eda-166">Method</span><span class="sxs-lookup"><span data-stu-id="c9eda-166">Method</span></span> |
| [<span data-ttu-id="c9eda-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="c9eda-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="c9eda-168">Method</span><span class="sxs-lookup"><span data-stu-id="c9eda-168">Method</span></span> |
| [<span data-ttu-id="c9eda-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="c9eda-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="c9eda-170">Method</span><span class="sxs-lookup"><span data-stu-id="c9eda-170">Method</span></span> |
| [<span data-ttu-id="c9eda-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="c9eda-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="c9eda-172">Method</span><span class="sxs-lookup"><span data-stu-id="c9eda-172">Method</span></span> |
| [<span data-ttu-id="c9eda-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="c9eda-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="c9eda-174">Method</span><span class="sxs-lookup"><span data-stu-id="c9eda-174">Method</span></span> |
| [<span data-ttu-id="c9eda-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="c9eda-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="c9eda-176">Method</span><span class="sxs-lookup"><span data-stu-id="c9eda-176">Method</span></span> |
| [<span data-ttu-id="c9eda-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="c9eda-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="c9eda-178">Method</span><span class="sxs-lookup"><span data-stu-id="c9eda-178">Method</span></span> |
| [<span data-ttu-id="c9eda-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="c9eda-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="c9eda-180">Method</span><span class="sxs-lookup"><span data-stu-id="c9eda-180">Method</span></span> |
| [<span data-ttu-id="c9eda-181">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c9eda-181">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="c9eda-182">Method</span><span class="sxs-lookup"><span data-stu-id="c9eda-182">Method</span></span> |
| [<span data-ttu-id="c9eda-183">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="c9eda-183">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="c9eda-184">Method</span><span class="sxs-lookup"><span data-stu-id="c9eda-184">Method</span></span> |
| [<span data-ttu-id="c9eda-185">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c9eda-185">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="c9eda-186">メソッド</span><span class="sxs-lookup"><span data-stu-id="c9eda-186">Method</span></span> |
| [<span data-ttu-id="c9eda-187">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c9eda-187">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="c9eda-188">メソッド</span><span class="sxs-lookup"><span data-stu-id="c9eda-188">Method</span></span> |

### <a name="example"></a><span data-ttu-id="c9eda-189">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-189">Example</span></span>

<span data-ttu-id="c9eda-190">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c9eda-190">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
};
```

### <a name="members"></a><span data-ttu-id="c9eda-191">Members</span><span class="sxs-lookup"><span data-stu-id="c9eda-191">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-12"></a><span data-ttu-id="c9eda-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="c9eda-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

<span data-ttu-id="c9eda-p103">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c9eda-195">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="c9eda-195">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="c9eda-196">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c9eda-196">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="c9eda-197">型</span><span class="sxs-lookup"><span data-stu-id="c9eda-197">Type</span></span>

*   <span data-ttu-id="c9eda-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="c9eda-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

##### <a name="requirements"></a><span data-ttu-id="c9eda-199">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-199">Requirements</span></span>

|<span data-ttu-id="c9eda-200">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-200">Requirement</span></span>| <span data-ttu-id="c9eda-201">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-201">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-202">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-202">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-203">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-203">1.0</span></span>|
|[<span data-ttu-id="c9eda-204">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-204">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-205">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-205">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-206">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-206">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-207">読み取り</span><span class="sxs-lookup"><span data-stu-id="c9eda-207">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9eda-208">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-208">Example</span></span>

<span data-ttu-id="c9eda-209">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-209">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
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

<br>

---
---

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="c9eda-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="c9eda-211">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-211">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="c9eda-212">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c9eda-212">Compose mode only.</span></span>

<span data-ttu-id="c9eda-213">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c9eda-213">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c9eda-214">ただし、Windows と Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-214">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="c9eda-215">500メンバーの最大数を取得します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-215">Get 500 members maximum.</span></span>
- <span data-ttu-id="c9eda-216">1回の呼び出しで最大100のメンバーを設定します。最大数は500メンバーです。</span><span class="sxs-lookup"><span data-stu-id="c9eda-216">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="c9eda-217">Type</span><span class="sxs-lookup"><span data-stu-id="c9eda-217">Type</span></span>

*   [<span data-ttu-id="c9eda-218">受信者</span><span class="sxs-lookup"><span data-stu-id="c9eda-218">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="c9eda-219">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9eda-219">Requirements</span></span>

|<span data-ttu-id="c9eda-220">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-220">Requirement</span></span>| <span data-ttu-id="c9eda-221">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-222">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-222">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-223">1.1</span><span class="sxs-lookup"><span data-stu-id="c9eda-223">1.1</span></span>|
|[<span data-ttu-id="c9eda-224">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-224">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-225">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-226">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-226">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-227">作成</span><span class="sxs-lookup"><span data-stu-id="c9eda-227">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c9eda-228">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-228">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

<br>

---
---

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-12"></a><span data-ttu-id="c9eda-229">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-229">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span></span>

<span data-ttu-id="c9eda-230">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-230">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c9eda-231">Type</span><span class="sxs-lookup"><span data-stu-id="c9eda-231">Type</span></span>

*   [<span data-ttu-id="c9eda-232">Body</span><span class="sxs-lookup"><span data-stu-id="c9eda-232">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="c9eda-233">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9eda-233">Requirements</span></span>

|<span data-ttu-id="c9eda-234">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-234">Requirement</span></span>| <span data-ttu-id="c9eda-235">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-236">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-237">1.1</span><span class="sxs-lookup"><span data-stu-id="c9eda-237">1.1</span></span>|
|[<span data-ttu-id="c9eda-238">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-238">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-239">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-240">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-240">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-241">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c9eda-241">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9eda-242">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-242">Example</span></span>

<span data-ttu-id="c9eda-243">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-243">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="c9eda-244">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="c9eda-244">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

<br>

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="c9eda-245">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-245">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="c9eda-246">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-246">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="c9eda-247">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c9eda-247">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c9eda-248">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c9eda-248">Read mode</span></span>

<span data-ttu-id="c9eda-249">`cc` プロパティは、メッセージの `EmailAddressDetails` 行にある各受信者について、\*\*\*\* オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-249">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="c9eda-250">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c9eda-250">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c9eda-251">ただし、Windows および Mac では、最大500メンバーを取得することができます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-251">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="c9eda-252">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c9eda-252">Compose mode</span></span>

<span data-ttu-id="c9eda-253">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="c9eda-254">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c9eda-254">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c9eda-255">ただし、Windows と Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-255">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="c9eda-256">500メンバーの最大数を取得します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-256">Get 500 members maximum.</span></span>
- <span data-ttu-id="c9eda-257">1回の呼び出しで最大100のメンバーを設定します。最大数は500メンバーです。</span><span class="sxs-lookup"><span data-stu-id="c9eda-257">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c9eda-258">Type</span><span class="sxs-lookup"><span data-stu-id="c9eda-258">Type</span></span>

*   <span data-ttu-id="c9eda-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9eda-260">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9eda-260">Requirements</span></span>

|<span data-ttu-id="c9eda-261">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-261">Requirement</span></span>| <span data-ttu-id="c9eda-262">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-263">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-264">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-264">1.0</span></span>|
|[<span data-ttu-id="c9eda-265">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-266">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-267">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-268">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c9eda-268">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="c9eda-269">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="c9eda-269">(nullable) conversationId: String</span></span>

<span data-ttu-id="c9eda-270">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="c9eda-p110">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p110">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="c9eda-p111">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p111">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="c9eda-275">Type</span><span class="sxs-lookup"><span data-stu-id="c9eda-275">Type</span></span>

*   <span data-ttu-id="c9eda-276">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9eda-277">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-277">Requirements</span></span>

|<span data-ttu-id="c9eda-278">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-278">Requirement</span></span>| <span data-ttu-id="c9eda-279">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-280">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-281">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-281">1.0</span></span>|
|[<span data-ttu-id="c9eda-282">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-283">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-284">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-285">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c9eda-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9eda-286">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-286">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="c9eda-287">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="c9eda-287">dateTimeCreated: Date</span></span>

<span data-ttu-id="c9eda-p112">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p112">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c9eda-290">型</span><span class="sxs-lookup"><span data-stu-id="c9eda-290">Type</span></span>

*   <span data-ttu-id="c9eda-291">日付</span><span class="sxs-lookup"><span data-stu-id="c9eda-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9eda-292">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-292">Requirements</span></span>

|<span data-ttu-id="c9eda-293">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-293">Requirement</span></span>| <span data-ttu-id="c9eda-294">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-295">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-296">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-296">1.0</span></span>|
|[<span data-ttu-id="c9eda-297">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-298">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-299">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-300">読み取り</span><span class="sxs-lookup"><span data-stu-id="c9eda-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9eda-301">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-301">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="c9eda-302">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="c9eda-302">dateTimeModified: Date</span></span>

<span data-ttu-id="c9eda-p113">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p113">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c9eda-305">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c9eda-305">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="c9eda-306">種類</span><span class="sxs-lookup"><span data-stu-id="c9eda-306">Type</span></span>

*   <span data-ttu-id="c9eda-307">日付</span><span class="sxs-lookup"><span data-stu-id="c9eda-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9eda-308">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-308">Requirements</span></span>

|<span data-ttu-id="c9eda-309">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-309">Requirement</span></span>| <span data-ttu-id="c9eda-310">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-312">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-312">1.0</span></span>|
|[<span data-ttu-id="c9eda-313">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-314">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-315">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-316">読み取り</span><span class="sxs-lookup"><span data-stu-id="c9eda-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9eda-317">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-317">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="c9eda-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="c9eda-319">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="c9eda-p114">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p114">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c9eda-322">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c9eda-322">Read mode</span></span>

<span data-ttu-id="c9eda-323">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-323">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="c9eda-324">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c9eda-324">Compose mode</span></span>

<span data-ttu-id="c9eda-325">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="c9eda-326">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c9eda-326">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="c9eda-327">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

##### <a name="type"></a><span data-ttu-id="c9eda-328">Type</span><span class="sxs-lookup"><span data-stu-id="c9eda-328">Type</span></span>

*   <span data-ttu-id="c9eda-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9eda-330">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9eda-330">Requirements</span></span>

|<span data-ttu-id="c9eda-331">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-331">Requirement</span></span>| <span data-ttu-id="c9eda-332">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-333">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-334">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-334">1.0</span></span>|
|[<span data-ttu-id="c9eda-335">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-336">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-337">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-338">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c9eda-338">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="c9eda-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="c9eda-p115">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p115">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="c9eda-p116">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p116">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c9eda-344">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="c9eda-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="c9eda-345">型</span><span class="sxs-lookup"><span data-stu-id="c9eda-345">Type</span></span>

*   [<span data-ttu-id="c9eda-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c9eda-346">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="c9eda-347">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9eda-347">Requirements</span></span>

|<span data-ttu-id="c9eda-348">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-348">Requirement</span></span>| <span data-ttu-id="c9eda-349">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-349">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-350">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-350">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-351">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-351">1.0</span></span>|
|[<span data-ttu-id="c9eda-352">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-352">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-353">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-354">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-354">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-355">読み取り</span><span class="sxs-lookup"><span data-stu-id="c9eda-355">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9eda-356">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-356">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="c9eda-357">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="c9eda-357">internetMessageId: String</span></span>

<span data-ttu-id="c9eda-p117">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p117">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c9eda-360">Type</span><span class="sxs-lookup"><span data-stu-id="c9eda-360">Type</span></span>

*   <span data-ttu-id="c9eda-361">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9eda-362">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-362">Requirements</span></span>

|<span data-ttu-id="c9eda-363">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-363">Requirement</span></span>| <span data-ttu-id="c9eda-364">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-365">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-366">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-366">1.0</span></span>|
|[<span data-ttu-id="c9eda-367">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-368">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-369">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-370">読み取り</span><span class="sxs-lookup"><span data-stu-id="c9eda-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9eda-371">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-371">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="c9eda-372">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="c9eda-372">itemClass: String</span></span>

<span data-ttu-id="c9eda-p118">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p118">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="c9eda-p119">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p119">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="c9eda-377">種類</span><span class="sxs-lookup"><span data-stu-id="c9eda-377">Type</span></span> | <span data-ttu-id="c9eda-378">説明</span><span class="sxs-lookup"><span data-stu-id="c9eda-378">Description</span></span> | <span data-ttu-id="c9eda-379">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="c9eda-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="c9eda-380">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="c9eda-380">Appointment items</span></span> | <span data-ttu-id="c9eda-381">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c9eda-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="c9eda-382">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="c9eda-382">Message items</span></span> | <span data-ttu-id="c9eda-383">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="c9eda-384">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="c9eda-385">Type</span><span class="sxs-lookup"><span data-stu-id="c9eda-385">Type</span></span>

*   <span data-ttu-id="c9eda-386">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9eda-387">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-387">Requirements</span></span>

|<span data-ttu-id="c9eda-388">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-388">Requirement</span></span>| <span data-ttu-id="c9eda-389">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-390">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-391">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-391">1.0</span></span>|
|[<span data-ttu-id="c9eda-392">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-393">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-394">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-395">読み取り</span><span class="sxs-lookup"><span data-stu-id="c9eda-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9eda-396">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-396">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="c9eda-397">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="c9eda-397">(nullable) itemId: String</span></span>

<span data-ttu-id="c9eda-p120">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p120">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c9eda-400">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="c9eda-400">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="c9eda-401">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="c9eda-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="c9eda-402">この値を使用して REST API を呼び出す前に、を`Office.context.mailbox.convertToRestId`使用して変換する必要があります。これは、要件セット1.3 から開始できます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-402">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="c9eda-403">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c9eda-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="c9eda-404">Type</span><span class="sxs-lookup"><span data-stu-id="c9eda-404">Type</span></span>

*   <span data-ttu-id="c9eda-405">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-405">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9eda-406">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-406">Requirements</span></span>

|<span data-ttu-id="c9eda-407">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-407">Requirement</span></span>| <span data-ttu-id="c9eda-408">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-409">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-410">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-410">1.0</span></span>|
|[<span data-ttu-id="c9eda-411">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-412">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-413">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-414">読み取り</span><span class="sxs-lookup"><span data-stu-id="c9eda-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9eda-415">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-415">Example</span></span>

<span data-ttu-id="c9eda-p122">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

<br>

---
---

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-12"></a><span data-ttu-id="c9eda-418">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-418">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span></span>

<span data-ttu-id="c9eda-419">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-419">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="c9eda-420">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="c9eda-420">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c9eda-421">型</span><span class="sxs-lookup"><span data-stu-id="c9eda-421">Type</span></span>

*   [<span data-ttu-id="c9eda-422">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="c9eda-422">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="c9eda-423">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-423">Requirements</span></span>

|<span data-ttu-id="c9eda-424">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-424">Requirement</span></span>| <span data-ttu-id="c9eda-425">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-426">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-427">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-427">1.0</span></span>|
|[<span data-ttu-id="c9eda-428">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-428">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-429">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-430">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-431">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c9eda-431">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9eda-432">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-432">Example</span></span>

```js
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

<br>

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-12"></a><span data-ttu-id="c9eda-433">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-433">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

<span data-ttu-id="c9eda-434">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-434">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c9eda-435">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c9eda-435">Read mode</span></span>

<span data-ttu-id="c9eda-436">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-436">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="c9eda-437">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c9eda-437">Compose mode</span></span>

<span data-ttu-id="c9eda-438">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-438">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c9eda-439">型</span><span class="sxs-lookup"><span data-stu-id="c9eda-439">Type</span></span>

*   <span data-ttu-id="c9eda-440">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-440">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9eda-441">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-441">Requirements</span></span>

|<span data-ttu-id="c9eda-442">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-442">Requirement</span></span>| <span data-ttu-id="c9eda-443">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-444">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-445">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-445">1.0</span></span>|
|[<span data-ttu-id="c9eda-446">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-447">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-447">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-448">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-449">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c9eda-449">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="c9eda-450">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="c9eda-450">normalizedSubject: String</span></span>

<span data-ttu-id="c9eda-p123">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="c9eda-p124">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="c9eda-455">Type</span><span class="sxs-lookup"><span data-stu-id="c9eda-455">Type</span></span>

*   <span data-ttu-id="c9eda-456">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-456">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9eda-457">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-457">Requirements</span></span>

|<span data-ttu-id="c9eda-458">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-458">Requirement</span></span>| <span data-ttu-id="c9eda-459">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-460">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-460">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-461">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-461">1.0</span></span>|
|[<span data-ttu-id="c9eda-462">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-462">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-463">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-464">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-464">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-465">読み取り</span><span class="sxs-lookup"><span data-stu-id="c9eda-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9eda-466">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-466">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="c9eda-467">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-467">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="c9eda-468">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-468">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="c9eda-469">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c9eda-469">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c9eda-470">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c9eda-470">Read mode</span></span>

<span data-ttu-id="c9eda-471">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-471">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="c9eda-472">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c9eda-472">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c9eda-473">ただし、Windows および Mac では、最大500メンバーを取得することができます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-473">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="c9eda-474">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c9eda-474">Compose mode</span></span>

<span data-ttu-id="c9eda-475">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-475">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="c9eda-476">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c9eda-476">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c9eda-477">ただし、Windows と Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-477">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="c9eda-478">500メンバーの最大数を取得します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-478">Get 500 members maximum.</span></span>
- <span data-ttu-id="c9eda-479">1回の呼び出しで最大100のメンバーを設定します。最大数は500メンバーです。</span><span class="sxs-lookup"><span data-stu-id="c9eda-479">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c9eda-480">Type</span><span class="sxs-lookup"><span data-stu-id="c9eda-480">Type</span></span>

*   <span data-ttu-id="c9eda-481">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-481">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9eda-482">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9eda-482">Requirements</span></span>

|<span data-ttu-id="c9eda-483">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-483">Requirement</span></span>| <span data-ttu-id="c9eda-484">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-484">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-485">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-485">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-486">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-486">1.0</span></span>|
|[<span data-ttu-id="c9eda-487">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-487">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-488">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-488">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-489">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-489">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-490">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c9eda-490">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="c9eda-491">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-491">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="c9eda-p128">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c9eda-494">型</span><span class="sxs-lookup"><span data-stu-id="c9eda-494">Type</span></span>

*   [<span data-ttu-id="c9eda-495">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c9eda-495">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="c9eda-496">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9eda-496">Requirements</span></span>

|<span data-ttu-id="c9eda-497">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-497">Requirement</span></span>| <span data-ttu-id="c9eda-498">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-498">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-499">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-499">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-500">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-500">1.0</span></span>|
|[<span data-ttu-id="c9eda-501">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-501">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-502">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-502">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-503">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-503">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-504">読み取り</span><span class="sxs-lookup"><span data-stu-id="c9eda-504">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9eda-505">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-505">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="c9eda-506">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-506">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="c9eda-507">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-507">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="c9eda-508">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c9eda-508">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c9eda-509">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c9eda-509">Read mode</span></span>

<span data-ttu-id="c9eda-510">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-510">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="c9eda-511">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c9eda-511">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c9eda-512">ただし、Windows および Mac では、最大500メンバーを取得することができます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-512">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="c9eda-513">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c9eda-513">Compose mode</span></span>

<span data-ttu-id="c9eda-514">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-514">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="c9eda-515">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c9eda-515">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c9eda-516">ただし、Windows と Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-516">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="c9eda-517">500メンバーの最大数を取得します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-517">Get 500 members maximum.</span></span>
- <span data-ttu-id="c9eda-518">1回の呼び出しで最大100のメンバーを設定します。最大数は500メンバーです。</span><span class="sxs-lookup"><span data-stu-id="c9eda-518">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="c9eda-519">Type</span><span class="sxs-lookup"><span data-stu-id="c9eda-519">Type</span></span>

*   <span data-ttu-id="c9eda-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9eda-521">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9eda-521">Requirements</span></span>

|<span data-ttu-id="c9eda-522">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-522">Requirement</span></span>| <span data-ttu-id="c9eda-523">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-524">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-525">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-525">1.0</span></span>|
|[<span data-ttu-id="c9eda-526">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-527">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-528">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-529">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c9eda-529">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="c9eda-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="c9eda-p132">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="c9eda-p133">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c9eda-535">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="c9eda-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="c9eda-536">型</span><span class="sxs-lookup"><span data-stu-id="c9eda-536">Type</span></span>

*   [<span data-ttu-id="c9eda-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c9eda-537">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="c9eda-538">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9eda-538">Requirements</span></span>

|<span data-ttu-id="c9eda-539">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-539">Requirement</span></span>| <span data-ttu-id="c9eda-540">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-541">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-542">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-542">1.0</span></span>|
|[<span data-ttu-id="c9eda-543">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-544">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-545">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-546">読み取り</span><span class="sxs-lookup"><span data-stu-id="c9eda-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9eda-547">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-547">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="c9eda-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="c9eda-549">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="c9eda-p134">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c9eda-552">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c9eda-552">Read mode</span></span>

<span data-ttu-id="c9eda-553">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-553">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="c9eda-554">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c9eda-554">Compose mode</span></span>

<span data-ttu-id="c9eda-555">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="c9eda-556">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c9eda-556">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="c9eda-557">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

##### <a name="type"></a><span data-ttu-id="c9eda-558">型</span><span class="sxs-lookup"><span data-stu-id="c9eda-558">Type</span></span>

*   <span data-ttu-id="c9eda-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9eda-560">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9eda-560">Requirements</span></span>

|<span data-ttu-id="c9eda-561">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-561">Requirement</span></span>| <span data-ttu-id="c9eda-562">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-563">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-564">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-564">1.0</span></span>|
|[<span data-ttu-id="c9eda-565">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-566">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-567">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-568">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c9eda-568">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-12"></a><span data-ttu-id="c9eda-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

<span data-ttu-id="c9eda-570">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="c9eda-571">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c9eda-572">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c9eda-572">Read mode</span></span>

<span data-ttu-id="c9eda-p136">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p136">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="c9eda-575">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c9eda-575">Compose mode</span></span>

<span data-ttu-id="c9eda-576">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="c9eda-577">型</span><span class="sxs-lookup"><span data-stu-id="c9eda-577">Type</span></span>

*   <span data-ttu-id="c9eda-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9eda-579">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9eda-579">Requirements</span></span>

|<span data-ttu-id="c9eda-580">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-580">Requirement</span></span>| <span data-ttu-id="c9eda-581">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-582">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-583">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-583">1.0</span></span>|
|[<span data-ttu-id="c9eda-584">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-585">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-586">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-587">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c9eda-587">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="c9eda-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="c9eda-589">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="c9eda-590">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c9eda-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c9eda-591">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c9eda-591">Read mode</span></span>

<span data-ttu-id="c9eda-592">`to` プロパティは、メッセージの `EmailAddressDetails` 行にある各受信者について、\*\*\*\* オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-592">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="c9eda-593">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c9eda-593">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c9eda-594">ただし、Windows および Mac では、最大500メンバーを取得することができます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-594">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="c9eda-595">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c9eda-595">Compose mode</span></span>

<span data-ttu-id="c9eda-596">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-596">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="c9eda-597">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c9eda-597">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c9eda-598">ただし、Windows と Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-598">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="c9eda-599">500メンバーの最大数を取得します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-599">Get 500 members maximum.</span></span>
- <span data-ttu-id="c9eda-600">1回の呼び出しで最大100のメンバーを設定します。最大数は500メンバーです。</span><span class="sxs-lookup"><span data-stu-id="c9eda-600">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c9eda-601">Type</span><span class="sxs-lookup"><span data-stu-id="c9eda-601">Type</span></span>

*   <span data-ttu-id="c9eda-602">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-602">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9eda-603">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9eda-603">Requirements</span></span>

|<span data-ttu-id="c9eda-604">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-604">Requirement</span></span>| <span data-ttu-id="c9eda-605">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-606">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-607">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-607">1.0</span></span>|
|[<span data-ttu-id="c9eda-608">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-608">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-609">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-609">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-610">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-610">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-611">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c9eda-611">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="c9eda-612">メソッド</span><span class="sxs-lookup"><span data-stu-id="c9eda-612">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="c9eda-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c9eda-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c9eda-614">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-614">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="c9eda-615">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-615">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="c9eda-616">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-616">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c9eda-617">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c9eda-617">Parameters</span></span>

|<span data-ttu-id="c9eda-618">名前</span><span class="sxs-lookup"><span data-stu-id="c9eda-618">Name</span></span>| <span data-ttu-id="c9eda-619">種類</span><span class="sxs-lookup"><span data-stu-id="c9eda-619">Type</span></span>| <span data-ttu-id="c9eda-620">属性</span><span class="sxs-lookup"><span data-stu-id="c9eda-620">Attributes</span></span>| <span data-ttu-id="c9eda-621">説明</span><span class="sxs-lookup"><span data-stu-id="c9eda-621">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="c9eda-622">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-622">String</span></span>||<span data-ttu-id="c9eda-p140">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p140">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="c9eda-625">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-625">String</span></span>||<span data-ttu-id="c9eda-p141">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p141">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="c9eda-628">Object</span><span class="sxs-lookup"><span data-stu-id="c9eda-628">Object</span></span>| <span data-ttu-id="c9eda-629">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-629">&lt;optional&gt;</span></span>|<span data-ttu-id="c9eda-630">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c9eda-630">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c9eda-631">Object</span><span class="sxs-lookup"><span data-stu-id="c9eda-631">Object</span></span>| <span data-ttu-id="c9eda-632">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-632">&lt;optional&gt;</span></span>|<span data-ttu-id="c9eda-633">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-633">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c9eda-634">function</span><span class="sxs-lookup"><span data-stu-id="c9eda-634">function</span></span>| <span data-ttu-id="c9eda-635">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-635">&lt;optional&gt;</span></span>|<span data-ttu-id="c9eda-636">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-636">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c9eda-637">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-637">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c9eda-638">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-638">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c9eda-639">エラー</span><span class="sxs-lookup"><span data-stu-id="c9eda-639">Errors</span></span>

| <span data-ttu-id="c9eda-640">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c9eda-640">Error code</span></span> | <span data-ttu-id="c9eda-641">説明</span><span class="sxs-lookup"><span data-stu-id="c9eda-641">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="c9eda-642">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="c9eda-642">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="c9eda-643">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="c9eda-643">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="c9eda-644">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-644">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c9eda-645">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-645">Requirements</span></span>

|<span data-ttu-id="c9eda-646">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-646">Requirement</span></span>| <span data-ttu-id="c9eda-647">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-647">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-648">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-648">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-649">1.1</span><span class="sxs-lookup"><span data-stu-id="c9eda-649">1.1</span></span>|
|[<span data-ttu-id="c9eda-650">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-650">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-651">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-651">ReadWriteItem</span></span>|
|[<span data-ttu-id="c9eda-652">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-652">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-653">作成</span><span class="sxs-lookup"><span data-stu-id="c9eda-653">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c9eda-654">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-654">Example</span></span>

```js
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

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="c9eda-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c9eda-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c9eda-656">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-656">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="c9eda-p142">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p142">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="c9eda-660">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-660">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="c9eda-661">Office アドインを Outlook on the web で実行している場合、編集中のアイテム以外のアイテムに `addItemAttachmentAsync` メソッドでアイテムを添付できます。ただし、これはサポートされていないため、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="c9eda-661">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c9eda-662">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c9eda-662">Parameters</span></span>

|<span data-ttu-id="c9eda-663">名前</span><span class="sxs-lookup"><span data-stu-id="c9eda-663">Name</span></span>| <span data-ttu-id="c9eda-664">種類</span><span class="sxs-lookup"><span data-stu-id="c9eda-664">Type</span></span>| <span data-ttu-id="c9eda-665">属性</span><span class="sxs-lookup"><span data-stu-id="c9eda-665">Attributes</span></span>| <span data-ttu-id="c9eda-666">説明</span><span class="sxs-lookup"><span data-stu-id="c9eda-666">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="c9eda-667">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-667">String</span></span>||<span data-ttu-id="c9eda-p143">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p143">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="c9eda-670">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-670">String</span></span>||<span data-ttu-id="c9eda-671">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="c9eda-671">The subject of the item to be attached.</span></span> <span data-ttu-id="c9eda-672">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c9eda-672">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="c9eda-673">Object</span><span class="sxs-lookup"><span data-stu-id="c9eda-673">Object</span></span>| <span data-ttu-id="c9eda-674">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-674">&lt;optional&gt;</span></span>|<span data-ttu-id="c9eda-675">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c9eda-675">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c9eda-676">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c9eda-676">Object</span></span>| <span data-ttu-id="c9eda-677">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-677">&lt;optional&gt;</span></span>|<span data-ttu-id="c9eda-678">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-678">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c9eda-679">関数</span><span class="sxs-lookup"><span data-stu-id="c9eda-679">function</span></span>| <span data-ttu-id="c9eda-680">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-680">&lt;optional&gt;</span></span>|<span data-ttu-id="c9eda-681">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-681">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c9eda-682">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-682">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c9eda-683">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-683">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c9eda-684">エラー</span><span class="sxs-lookup"><span data-stu-id="c9eda-684">Errors</span></span>

| <span data-ttu-id="c9eda-685">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c9eda-685">Error code</span></span> | <span data-ttu-id="c9eda-686">説明</span><span class="sxs-lookup"><span data-stu-id="c9eda-686">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="c9eda-687">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-687">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c9eda-688">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9eda-688">Requirements</span></span>

|<span data-ttu-id="c9eda-689">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-689">Requirement</span></span>| <span data-ttu-id="c9eda-690">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-691">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-692">1.1</span><span class="sxs-lookup"><span data-stu-id="c9eda-692">1.1</span></span>|
|[<span data-ttu-id="c9eda-693">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-693">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-694">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-694">ReadWriteItem</span></span>|
|[<span data-ttu-id="c9eda-695">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-695">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-696">作成</span><span class="sxs-lookup"><span data-stu-id="c9eda-696">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c9eda-697">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-697">Example</span></span>

<span data-ttu-id="c9eda-698">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-698">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
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

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="c9eda-699">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="c9eda-699">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="c9eda-700">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-700">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c9eda-701">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c9eda-701">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c9eda-702">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-702">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c9eda-703">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="c9eda-703">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="c9eda-p145">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c9eda-707">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c9eda-707">Parameters</span></span>

|<span data-ttu-id="c9eda-708">名前</span><span class="sxs-lookup"><span data-stu-id="c9eda-708">Name</span></span>| <span data-ttu-id="c9eda-709">種類</span><span class="sxs-lookup"><span data-stu-id="c9eda-709">Type</span></span>| <span data-ttu-id="c9eda-710">説明</span><span class="sxs-lookup"><span data-stu-id="c9eda-710">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="c9eda-711">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="c9eda-711">String &#124; Object</span></span>| |<span data-ttu-id="c9eda-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c9eda-714">**または**</span><span class="sxs-lookup"><span data-stu-id="c9eda-714">**OR**</span></span><br/><span data-ttu-id="c9eda-p147">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="c9eda-717">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-717">String</span></span> | <span data-ttu-id="c9eda-718">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-718">&lt;optional&gt;</span></span> | <span data-ttu-id="c9eda-p148">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="c9eda-721">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-721">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="c9eda-722">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-722">&lt;optional&gt;</span></span> | <span data-ttu-id="c9eda-723">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="c9eda-723">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="c9eda-724">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-724">String</span></span> | | <span data-ttu-id="c9eda-p149">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="c9eda-727">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-727">String</span></span> | | <span data-ttu-id="c9eda-728">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c9eda-728">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="c9eda-729">文字列</span><span class="sxs-lookup"><span data-stu-id="c9eda-729">String</span></span> | | <span data-ttu-id="c9eda-p150">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="c9eda-732">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-732">String</span></span> | | <span data-ttu-id="c9eda-p151">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="c9eda-736">function</span><span class="sxs-lookup"><span data-stu-id="c9eda-736">function</span></span> | <span data-ttu-id="c9eda-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-737">&lt;optional&gt;</span></span> | <span data-ttu-id="c9eda-738">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-738">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c9eda-739">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-739">Requirements</span></span>

|<span data-ttu-id="c9eda-740">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-740">Requirement</span></span>| <span data-ttu-id="c9eda-741">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-741">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-742">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-742">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-743">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-743">1.0</span></span>|
|[<span data-ttu-id="c9eda-744">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-744">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-745">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-745">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-746">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-746">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-747">読み取り</span><span class="sxs-lookup"><span data-stu-id="c9eda-747">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c9eda-748">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-748">Examples</span></span>

<span data-ttu-id="c9eda-749">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-749">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="c9eda-750">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-750">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="c9eda-751">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-751">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c9eda-752">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-752">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c9eda-753">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-753">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c9eda-754">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-754">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

<br>

---
---

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="c9eda-755">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="c9eda-755">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="c9eda-756">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-756">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c9eda-757">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c9eda-757">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c9eda-758">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-758">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c9eda-759">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="c9eda-759">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="c9eda-p152">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p152">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c9eda-763">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c9eda-763">Parameters</span></span>

|<span data-ttu-id="c9eda-764">名前</span><span class="sxs-lookup"><span data-stu-id="c9eda-764">Name</span></span>| <span data-ttu-id="c9eda-765">種類</span><span class="sxs-lookup"><span data-stu-id="c9eda-765">Type</span></span>| <span data-ttu-id="c9eda-766">説明</span><span class="sxs-lookup"><span data-stu-id="c9eda-766">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="c9eda-767">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="c9eda-767">String &#124; Object</span></span>| | <span data-ttu-id="c9eda-p153">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c9eda-770">**または**</span><span class="sxs-lookup"><span data-stu-id="c9eda-770">**OR**</span></span><br/><span data-ttu-id="c9eda-p154">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p154">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="c9eda-773">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-773">String</span></span> | <span data-ttu-id="c9eda-774">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-774">&lt;optional&gt;</span></span> | <span data-ttu-id="c9eda-p155">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p155">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="c9eda-777">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-777">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="c9eda-778">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-778">&lt;optional&gt;</span></span> | <span data-ttu-id="c9eda-779">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="c9eda-779">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="c9eda-780">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-780">String</span></span> | | <span data-ttu-id="c9eda-p156">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p156">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="c9eda-783">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-783">String</span></span> | | <span data-ttu-id="c9eda-784">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c9eda-784">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="c9eda-785">文字列</span><span class="sxs-lookup"><span data-stu-id="c9eda-785">String</span></span> | | <span data-ttu-id="c9eda-p157">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p157">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="c9eda-788">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-788">String</span></span> | | <span data-ttu-id="c9eda-p158">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="c9eda-792">function</span><span class="sxs-lookup"><span data-stu-id="c9eda-792">function</span></span> | <span data-ttu-id="c9eda-793">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-793">&lt;optional&gt;</span></span> | <span data-ttu-id="c9eda-794">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-794">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c9eda-795">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9eda-795">Requirements</span></span>

|<span data-ttu-id="c9eda-796">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-796">Requirement</span></span>| <span data-ttu-id="c9eda-797">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-797">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-798">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-798">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-799">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-799">1.0</span></span>|
|[<span data-ttu-id="c9eda-800">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-800">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-801">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-801">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-802">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-802">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-803">読み取り</span><span class="sxs-lookup"><span data-stu-id="c9eda-803">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c9eda-804">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-804">Examples</span></span>

<span data-ttu-id="c9eda-805">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-805">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="c9eda-806">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-806">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="c9eda-807">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-807">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c9eda-808">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-808">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c9eda-809">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-809">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c9eda-810">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-810">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-12"></a><span data-ttu-id="c9eda-811">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="c9eda-811">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="c9eda-812">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-812">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="c9eda-813">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c9eda-813">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9eda-814">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-814">Requirements</span></span>

|<span data-ttu-id="c9eda-815">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-815">Requirement</span></span>| <span data-ttu-id="c9eda-816">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-817">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-818">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-818">1.0</span></span>|
|[<span data-ttu-id="c9eda-819">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-820">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-820">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-821">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-822">読み取り</span><span class="sxs-lookup"><span data-stu-id="c9eda-822">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c9eda-823">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c9eda-823">Returns:</span></span>

<span data-ttu-id="c9eda-824">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="c9eda-824">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span></span>

##### <a name="example"></a><span data-ttu-id="c9eda-825">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-825">Example</span></span>

<span data-ttu-id="c9eda-826">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="c9eda-826">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="c9eda-827">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="c9eda-827">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="c9eda-828">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-828">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="c9eda-829">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c9eda-829">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c9eda-830">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c9eda-830">Parameters</span></span>

|<span data-ttu-id="c9eda-831">名前</span><span class="sxs-lookup"><span data-stu-id="c9eda-831">Name</span></span>| <span data-ttu-id="c9eda-832">種類</span><span class="sxs-lookup"><span data-stu-id="c9eda-832">Type</span></span>| <span data-ttu-id="c9eda-833">説明</span><span class="sxs-lookup"><span data-stu-id="c9eda-833">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="c9eda-834">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="c9eda-834">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.2)|<span data-ttu-id="c9eda-835">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="c9eda-835">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c9eda-836">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9eda-836">Requirements</span></span>

|<span data-ttu-id="c9eda-837">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-837">Requirement</span></span>| <span data-ttu-id="c9eda-838">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-839">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-840">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-840">1.0</span></span>|
|[<span data-ttu-id="c9eda-841">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-841">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-842">制限あり</span><span class="sxs-lookup"><span data-stu-id="c9eda-842">Restricted</span></span>|
|[<span data-ttu-id="c9eda-843">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-843">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-844">読み取り</span><span class="sxs-lookup"><span data-stu-id="c9eda-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c9eda-845">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c9eda-845">Returns:</span></span>

<span data-ttu-id="c9eda-846">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-846">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="c9eda-847">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-847">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="c9eda-848">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="c9eda-848">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="c9eda-849">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="c9eda-849">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="c9eda-850">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="c9eda-850">Value of `entityType`</span></span> | <span data-ttu-id="c9eda-851">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="c9eda-851">Type of objects in returned array</span></span> | <span data-ttu-id="c9eda-852">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-852">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="c9eda-853">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-853">String</span></span> | <span data-ttu-id="c9eda-854">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="c9eda-854">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="c9eda-855">連絡先</span><span class="sxs-lookup"><span data-stu-id="c9eda-855">Contact</span></span> | <span data-ttu-id="c9eda-856">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c9eda-856">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="c9eda-857">文字列</span><span class="sxs-lookup"><span data-stu-id="c9eda-857">String</span></span> | <span data-ttu-id="c9eda-858">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c9eda-858">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="c9eda-859">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="c9eda-859">MeetingSuggestion</span></span> | <span data-ttu-id="c9eda-860">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c9eda-860">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="c9eda-861">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="c9eda-861">PhoneNumber</span></span> | <span data-ttu-id="c9eda-862">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="c9eda-862">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="c9eda-863">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="c9eda-863">TaskSuggestion</span></span> | <span data-ttu-id="c9eda-864">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c9eda-864">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="c9eda-865">文字列</span><span class="sxs-lookup"><span data-stu-id="c9eda-865">String</span></span> | <span data-ttu-id="c9eda-866">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="c9eda-866">**Restricted**</span></span> |

<span data-ttu-id="c9eda-867">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="c9eda-867">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

##### <a name="example"></a><span data-ttu-id="c9eda-868">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-868">Example</span></span>

<span data-ttu-id="c9eda-869">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-869">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="c9eda-870">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="c9eda-870">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="c9eda-871">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-871">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c9eda-872">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c9eda-872">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c9eda-873">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-873">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c9eda-874">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c9eda-874">Parameters</span></span>

|<span data-ttu-id="c9eda-875">名前</span><span class="sxs-lookup"><span data-stu-id="c9eda-875">Name</span></span>| <span data-ttu-id="c9eda-876">種類</span><span class="sxs-lookup"><span data-stu-id="c9eda-876">Type</span></span>| <span data-ttu-id="c9eda-877">説明</span><span class="sxs-lookup"><span data-stu-id="c9eda-877">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="c9eda-878">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-878">String</span></span>|<span data-ttu-id="c9eda-879">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="c9eda-879">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c9eda-880">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9eda-880">Requirements</span></span>

|<span data-ttu-id="c9eda-881">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-881">Requirement</span></span>| <span data-ttu-id="c9eda-882">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-882">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-883">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-883">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-884">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-884">1.0</span></span>|
|[<span data-ttu-id="c9eda-885">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-885">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-886">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-886">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-887">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-887">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-888">読み取り</span><span class="sxs-lookup"><span data-stu-id="c9eda-888">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c9eda-889">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c9eda-889">Returns:</span></span>

<span data-ttu-id="c9eda-p160">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="c9eda-892">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="c9eda-892">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="c9eda-893">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="c9eda-893">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="c9eda-894">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-894">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c9eda-895">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c9eda-895">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c9eda-p161">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c9eda-899">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="c9eda-899">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c9eda-900">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="c9eda-900">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="c9eda-p162">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9eda-903">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9eda-903">Requirements</span></span>

|<span data-ttu-id="c9eda-904">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-904">Requirement</span></span>| <span data-ttu-id="c9eda-905">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-906">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-907">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-907">1.0</span></span>|
|[<span data-ttu-id="c9eda-908">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-908">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-909">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-910">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-910">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-911">読み取り</span><span class="sxs-lookup"><span data-stu-id="c9eda-911">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c9eda-912">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c9eda-912">Returns:</span></span>

<span data-ttu-id="c9eda-p163">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="c9eda-915">型: Object</span><span class="sxs-lookup"><span data-stu-id="c9eda-915">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="c9eda-916">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-916">Example</span></span>

<span data-ttu-id="c9eda-917">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="c9eda-917">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="c9eda-918">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="c9eda-918">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="c9eda-919">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-919">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c9eda-920">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c9eda-920">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c9eda-921">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="c9eda-921">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="c9eda-p164">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c9eda-924">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c9eda-924">Parameters</span></span>

|<span data-ttu-id="c9eda-925">名前</span><span class="sxs-lookup"><span data-stu-id="c9eda-925">Name</span></span>| <span data-ttu-id="c9eda-926">種類</span><span class="sxs-lookup"><span data-stu-id="c9eda-926">Type</span></span>| <span data-ttu-id="c9eda-927">説明</span><span class="sxs-lookup"><span data-stu-id="c9eda-927">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="c9eda-928">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-928">String</span></span>|<span data-ttu-id="c9eda-929">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="c9eda-929">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c9eda-930">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9eda-930">Requirements</span></span>

|<span data-ttu-id="c9eda-931">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-931">Requirement</span></span>| <span data-ttu-id="c9eda-932">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-933">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-934">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-934">1.0</span></span>|
|[<span data-ttu-id="c9eda-935">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-936">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-937">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-938">読み取り</span><span class="sxs-lookup"><span data-stu-id="c9eda-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c9eda-939">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c9eda-939">Returns:</span></span>

<span data-ttu-id="c9eda-940">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="c9eda-940">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="c9eda-941">型: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="c9eda-941">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="c9eda-942">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-942">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="c9eda-943">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="c9eda-943">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="c9eda-944">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-944">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="c9eda-p165">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c9eda-947">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c9eda-947">Parameters</span></span>

|<span data-ttu-id="c9eda-948">名前</span><span class="sxs-lookup"><span data-stu-id="c9eda-948">Name</span></span>| <span data-ttu-id="c9eda-949">種類</span><span class="sxs-lookup"><span data-stu-id="c9eda-949">Type</span></span>| <span data-ttu-id="c9eda-950">属性</span><span class="sxs-lookup"><span data-stu-id="c9eda-950">Attributes</span></span>| <span data-ttu-id="c9eda-951">説明</span><span class="sxs-lookup"><span data-stu-id="c9eda-951">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="c9eda-952">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c9eda-952">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="c9eda-p166">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="c9eda-956">Object</span><span class="sxs-lookup"><span data-stu-id="c9eda-956">Object</span></span>| <span data-ttu-id="c9eda-957">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-957">&lt;optional&gt;</span></span>|<span data-ttu-id="c9eda-958">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c9eda-958">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c9eda-959">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c9eda-959">Object</span></span>| <span data-ttu-id="c9eda-960">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-960">&lt;optional&gt;</span></span>|<span data-ttu-id="c9eda-961">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-961">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c9eda-962">function</span><span class="sxs-lookup"><span data-stu-id="c9eda-962">function</span></span>||<span data-ttu-id="c9eda-963">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-963">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c9eda-964">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-964">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="c9eda-965">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="c9eda-965">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c9eda-966">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9eda-966">Requirements</span></span>

|<span data-ttu-id="c9eda-967">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-967">Requirement</span></span>| <span data-ttu-id="c9eda-968">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-968">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-969">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-969">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-970">1.2</span><span class="sxs-lookup"><span data-stu-id="c9eda-970">1.2</span></span>|
|[<span data-ttu-id="c9eda-971">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-971">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-972">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-972">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-973">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-973">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-974">作成</span><span class="sxs-lookup"><span data-stu-id="c9eda-974">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="c9eda-975">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c9eda-975">Returns:</span></span>

<span data-ttu-id="c9eda-976">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="c9eda-976">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="c9eda-977">型:String</span><span class="sxs-lookup"><span data-stu-id="c9eda-977">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="c9eda-978">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-978">Example</span></span>

```js
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

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="c9eda-979">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c9eda-979">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="c9eda-980">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-980">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="c9eda-p168">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c9eda-984">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c9eda-984">Parameters</span></span>

|<span data-ttu-id="c9eda-985">名前</span><span class="sxs-lookup"><span data-stu-id="c9eda-985">Name</span></span>| <span data-ttu-id="c9eda-986">種類</span><span class="sxs-lookup"><span data-stu-id="c9eda-986">Type</span></span>| <span data-ttu-id="c9eda-987">属性</span><span class="sxs-lookup"><span data-stu-id="c9eda-987">Attributes</span></span>| <span data-ttu-id="c9eda-988">説明</span><span class="sxs-lookup"><span data-stu-id="c9eda-988">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="c9eda-989">function</span><span class="sxs-lookup"><span data-stu-id="c9eda-989">function</span></span>||<span data-ttu-id="c9eda-990">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-990">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c9eda-991">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-991">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="c9eda-992">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-992">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="c9eda-993">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c9eda-993">Object</span></span>| <span data-ttu-id="c9eda-994">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-994">&lt;optional&gt;</span></span>|<span data-ttu-id="c9eda-995">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-995">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="c9eda-996">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-996">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c9eda-997">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-997">Requirements</span></span>

|<span data-ttu-id="c9eda-998">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-998">Requirement</span></span>| <span data-ttu-id="c9eda-999">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-999">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-1000">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-1000">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-1001">1.0</span><span class="sxs-lookup"><span data-stu-id="c9eda-1001">1.0</span></span>|
|[<span data-ttu-id="c9eda-1002">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-1002">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-1003">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-1003">ReadItem</span></span>|
|[<span data-ttu-id="c9eda-1004">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-1004">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-1005">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c9eda-1005">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9eda-1006">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-1006">Example</span></span>

<span data-ttu-id="c9eda-p171">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

<br>

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="c9eda-1010">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c9eda-1010">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="c9eda-1011">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-1011">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="c9eda-1012">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-1012">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="c9eda-1013">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="c9eda-1013">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="c9eda-1014">Outlook on the web とモバイル デバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="c9eda-1014">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="c9eda-1015">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-1015">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c9eda-1016">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c9eda-1016">Parameters</span></span>

|<span data-ttu-id="c9eda-1017">名前</span><span class="sxs-lookup"><span data-stu-id="c9eda-1017">Name</span></span>| <span data-ttu-id="c9eda-1018">種類</span><span class="sxs-lookup"><span data-stu-id="c9eda-1018">Type</span></span>| <span data-ttu-id="c9eda-1019">属性</span><span class="sxs-lookup"><span data-stu-id="c9eda-1019">Attributes</span></span>| <span data-ttu-id="c9eda-1020">説明</span><span class="sxs-lookup"><span data-stu-id="c9eda-1020">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="c9eda-1021">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-1021">String</span></span>||<span data-ttu-id="c9eda-1022">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="c9eda-1022">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="c9eda-1023">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c9eda-1023">Object</span></span>| <span data-ttu-id="c9eda-1024">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-1024">&lt;optional&gt;</span></span>|<span data-ttu-id="c9eda-1025">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c9eda-1025">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c9eda-1026">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c9eda-1026">Object</span></span>| <span data-ttu-id="c9eda-1027">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-1027">&lt;optional&gt;</span></span>|<span data-ttu-id="c9eda-1028">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-1028">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c9eda-1029">関数</span><span class="sxs-lookup"><span data-stu-id="c9eda-1029">function</span></span>| <span data-ttu-id="c9eda-1030">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-1030">&lt;optional&gt;</span></span>|<span data-ttu-id="c9eda-1031">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-1031">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c9eda-1032">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-1032">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c9eda-1033">エラー</span><span class="sxs-lookup"><span data-stu-id="c9eda-1033">Errors</span></span>

| <span data-ttu-id="c9eda-1034">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c9eda-1034">Error code</span></span> | <span data-ttu-id="c9eda-1035">説明</span><span class="sxs-lookup"><span data-stu-id="c9eda-1035">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="c9eda-1036">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="c9eda-1036">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c9eda-1037">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9eda-1037">Requirements</span></span>

|<span data-ttu-id="c9eda-1038">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-1038">Requirement</span></span>| <span data-ttu-id="c9eda-1039">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-1039">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-1040">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-1040">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-1041">1.1</span><span class="sxs-lookup"><span data-stu-id="c9eda-1041">1.1</span></span>|
|[<span data-ttu-id="c9eda-1042">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-1042">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-1043">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-1043">ReadWriteItem</span></span>|
|[<span data-ttu-id="c9eda-1044">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-1044">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-1045">作成</span><span class="sxs-lookup"><span data-stu-id="c9eda-1045">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c9eda-1046">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-1046">Example</span></span>

<span data-ttu-id="c9eda-1047">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-1047">The following code removes an attachment with an identifier of '0'.</span></span>

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

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="c9eda-1048">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="c9eda-1048">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="c9eda-1049">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="c9eda-1049">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="c9eda-p173">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p173">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c9eda-1053">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c9eda-1053">Parameters</span></span>

|<span data-ttu-id="c9eda-1054">名前</span><span class="sxs-lookup"><span data-stu-id="c9eda-1054">Name</span></span>| <span data-ttu-id="c9eda-1055">種類</span><span class="sxs-lookup"><span data-stu-id="c9eda-1055">Type</span></span>| <span data-ttu-id="c9eda-1056">属性</span><span class="sxs-lookup"><span data-stu-id="c9eda-1056">Attributes</span></span>| <span data-ttu-id="c9eda-1057">説明</span><span class="sxs-lookup"><span data-stu-id="c9eda-1057">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="c9eda-1058">String</span><span class="sxs-lookup"><span data-stu-id="c9eda-1058">String</span></span>||<span data-ttu-id="c9eda-p174">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-p174">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="c9eda-1062">Object</span><span class="sxs-lookup"><span data-stu-id="c9eda-1062">Object</span></span>| <span data-ttu-id="c9eda-1063">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-1063">&lt;optional&gt;</span></span>|<span data-ttu-id="c9eda-1064">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c9eda-1064">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c9eda-1065">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c9eda-1065">Object</span></span>| <span data-ttu-id="c9eda-1066">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-1066">&lt;optional&gt;</span></span>|<span data-ttu-id="c9eda-1067">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-1067">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="c9eda-1068">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c9eda-1068">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="c9eda-1069">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c9eda-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="c9eda-1070">`text` の場合、Outlook on the web とデスクトップ クライアントでは現在のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-1070">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="c9eda-1071">フィールドが HTML エディターの場合、データが HTML の場合でもテキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-1071">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="c9eda-1072">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Outlook on the web では現在のスタイルが適用され、Outlook デスクトップ クライアントでは既定のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-1072">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="c9eda-1073">フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-1073">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="c9eda-1074">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="c9eda-1074">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="c9eda-1075">function</span><span class="sxs-lookup"><span data-stu-id="c9eda-1075">function</span></span>||<span data-ttu-id="c9eda-1076">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="c9eda-1076">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c9eda-1077">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-1077">Requirements</span></span>

|<span data-ttu-id="c9eda-1078">要件</span><span class="sxs-lookup"><span data-stu-id="c9eda-1078">Requirement</span></span>| <span data-ttu-id="c9eda-1079">値</span><span class="sxs-lookup"><span data-stu-id="c9eda-1079">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9eda-1080">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9eda-1080">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9eda-1081">1.2</span><span class="sxs-lookup"><span data-stu-id="c9eda-1081">1.2</span></span>|
|[<span data-ttu-id="c9eda-1082">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9eda-1082">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9eda-1083">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c9eda-1083">ReadWriteItem</span></span>|
|[<span data-ttu-id="c9eda-1084">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9eda-1084">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9eda-1085">作成</span><span class="sxs-lookup"><span data-stu-id="c9eda-1085">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c9eda-1086">例</span><span class="sxs-lookup"><span data-stu-id="c9eda-1086">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
