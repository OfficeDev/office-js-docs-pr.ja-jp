---
title: Office. メールボックス-要件セット1.6
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: cc7897f791c5a07ed5c17a686b6601a1a7633f00
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838544"
---
# <a name="item"></a><span data-ttu-id="1a9ce-102">item</span><span class="sxs-lookup"><span data-stu-id="1a9ce-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="1a9ce-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="1a9ce-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="1a9ce-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-106">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-106">Requirements</span></span>

|<span data-ttu-id="1a9ce-107">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-107">Requirement</span></span>| <span data-ttu-id="1a9ce-108">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-110">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-110">1.0</span></span>|
|[<span data-ttu-id="1a9ce-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="1a9ce-112">Restricted</span></span>|
|[<span data-ttu-id="1a9ce-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1a9ce-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1a9ce-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="1a9ce-115">Members and methods</span></span>

| <span data-ttu-id="1a9ce-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-116">Member</span></span> | <span data-ttu-id="1a9ce-117">種類</span><span class="sxs-lookup"><span data-stu-id="1a9ce-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1a9ce-118">attachments</span><span class="sxs-lookup"><span data-stu-id="1a9ce-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="1a9ce-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-119">Member</span></span> |
| [<span data-ttu-id="1a9ce-120">bcc</span><span class="sxs-lookup"><span data-stu-id="1a9ce-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="1a9ce-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-121">Member</span></span> |
| [<span data-ttu-id="1a9ce-122">body</span><span class="sxs-lookup"><span data-stu-id="1a9ce-122">body</span></span>](#body-body) | <span data-ttu-id="1a9ce-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-123">Member</span></span> |
| [<span data-ttu-id="1a9ce-124">cc</span><span class="sxs-lookup"><span data-stu-id="1a9ce-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="1a9ce-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-125">Member</span></span> |
| [<span data-ttu-id="1a9ce-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="1a9ce-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="1a9ce-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-127">Member</span></span> |
| [<span data-ttu-id="1a9ce-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="1a9ce-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="1a9ce-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-129">Member</span></span> |
| [<span data-ttu-id="1a9ce-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="1a9ce-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="1a9ce-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-131">Member</span></span> |
| [<span data-ttu-id="1a9ce-132">end</span><span class="sxs-lookup"><span data-stu-id="1a9ce-132">end</span></span>](#end-datetime) | <span data-ttu-id="1a9ce-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-133">Member</span></span> |
| [<span data-ttu-id="1a9ce-134">from</span><span class="sxs-lookup"><span data-stu-id="1a9ce-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="1a9ce-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-135">Member</span></span> |
| [<span data-ttu-id="1a9ce-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="1a9ce-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="1a9ce-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-137">Member</span></span> |
| [<span data-ttu-id="1a9ce-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="1a9ce-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="1a9ce-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-139">Member</span></span> |
| [<span data-ttu-id="1a9ce-140">itemId</span><span class="sxs-lookup"><span data-stu-id="1a9ce-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="1a9ce-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-141">Member</span></span> |
| [<span data-ttu-id="1a9ce-142">itemType</span><span class="sxs-lookup"><span data-stu-id="1a9ce-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="1a9ce-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-143">Member</span></span> |
| [<span data-ttu-id="1a9ce-144">location</span><span class="sxs-lookup"><span data-stu-id="1a9ce-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="1a9ce-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-145">Member</span></span> |
| [<span data-ttu-id="1a9ce-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="1a9ce-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="1a9ce-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-147">Member</span></span> |
| [<span data-ttu-id="1a9ce-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="1a9ce-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="1a9ce-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-149">Member</span></span> |
| [<span data-ttu-id="1a9ce-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="1a9ce-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="1a9ce-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-151">Member</span></span> |
| [<span data-ttu-id="1a9ce-152">organizer</span><span class="sxs-lookup"><span data-stu-id="1a9ce-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="1a9ce-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-153">Member</span></span> |
| [<span data-ttu-id="1a9ce-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="1a9ce-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="1a9ce-155">Member</span><span class="sxs-lookup"><span data-stu-id="1a9ce-155">Member</span></span> |
| [<span data-ttu-id="1a9ce-156">sender</span><span class="sxs-lookup"><span data-stu-id="1a9ce-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="1a9ce-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-157">Member</span></span> |
| [<span data-ttu-id="1a9ce-158">start</span><span class="sxs-lookup"><span data-stu-id="1a9ce-158">start</span></span>](#start-datetime) | <span data-ttu-id="1a9ce-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-159">Member</span></span> |
| [<span data-ttu-id="1a9ce-160">subject</span><span class="sxs-lookup"><span data-stu-id="1a9ce-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="1a9ce-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-161">Member</span></span> |
| [<span data-ttu-id="1a9ce-162">to</span><span class="sxs-lookup"><span data-stu-id="1a9ce-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="1a9ce-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-163">Member</span></span> |
| [<span data-ttu-id="1a9ce-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1a9ce-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="1a9ce-165">メソッド</span><span class="sxs-lookup"><span data-stu-id="1a9ce-165">Method</span></span> |
| [<span data-ttu-id="1a9ce-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1a9ce-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="1a9ce-167">メソッド</span><span class="sxs-lookup"><span data-stu-id="1a9ce-167">Method</span></span> |
| [<span data-ttu-id="1a9ce-168">close</span><span class="sxs-lookup"><span data-stu-id="1a9ce-168">close</span></span>](#close) | <span data-ttu-id="1a9ce-169">メソッド</span><span class="sxs-lookup"><span data-stu-id="1a9ce-169">Method</span></span> |
| [<span data-ttu-id="1a9ce-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="1a9ce-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="1a9ce-171">メソッド</span><span class="sxs-lookup"><span data-stu-id="1a9ce-171">Method</span></span> |
| [<span data-ttu-id="1a9ce-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="1a9ce-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="1a9ce-173">メソッド</span><span class="sxs-lookup"><span data-stu-id="1a9ce-173">Method</span></span> |
| [<span data-ttu-id="1a9ce-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="1a9ce-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="1a9ce-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="1a9ce-175">Method</span></span> |
| [<span data-ttu-id="1a9ce-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="1a9ce-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="1a9ce-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="1a9ce-177">Method</span></span> |
| [<span data-ttu-id="1a9ce-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="1a9ce-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="1a9ce-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="1a9ce-179">Method</span></span> |
| [<span data-ttu-id="1a9ce-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="1a9ce-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="1a9ce-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="1a9ce-181">Method</span></span> |
| [<span data-ttu-id="1a9ce-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="1a9ce-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="1a9ce-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="1a9ce-183">Method</span></span> |
| [<span data-ttu-id="1a9ce-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="1a9ce-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="1a9ce-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="1a9ce-185">Method</span></span> |
| [<span data-ttu-id="1a9ce-186">office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="1a9ce-186">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="1a9ce-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="1a9ce-187">Method</span></span> |
| [<span data-ttu-id="1a9ce-188">office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="1a9ce-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="1a9ce-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="1a9ce-189">Method</span></span> |
| [<span data-ttu-id="1a9ce-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="1a9ce-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="1a9ce-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="1a9ce-191">Method</span></span> |
| [<span data-ttu-id="1a9ce-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1a9ce-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="1a9ce-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="1a9ce-193">Method</span></span> |
| [<span data-ttu-id="1a9ce-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="1a9ce-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="1a9ce-195">メソッド</span><span class="sxs-lookup"><span data-stu-id="1a9ce-195">Method</span></span> |
| [<span data-ttu-id="1a9ce-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="1a9ce-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="1a9ce-197">メソッド</span><span class="sxs-lookup"><span data-stu-id="1a9ce-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="1a9ce-198">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-198">Example</span></span>

<span data-ttu-id="1a9ce-199">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="1a9ce-200">メンバー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="1a9ce-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="1a9ce-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="1a9ce-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1a9ce-204">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="1a9ce-205">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="1a9ce-206">タイプ</span><span class="sxs-lookup"><span data-stu-id="1a9ce-206">Type</span></span>

*   <span data-ttu-id="1a9ce-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="1a9ce-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-208">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-208">Requirements</span></span>

|<span data-ttu-id="1a9ce-209">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-209">Requirement</span></span>| <span data-ttu-id="1a9ce-210">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-211">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-212">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-212">1.0</span></span>|
|[<span data-ttu-id="1a9ce-213">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-213">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-214">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-215">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-216">読み取り</span><span class="sxs-lookup"><span data-stu-id="1a9ce-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1a9ce-217">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-217">Example</span></span>

<span data-ttu-id="1a9ce-218">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="1a9ce-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="1a9ce-220">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="1a9ce-221">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-221">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1a9ce-222">タイプ</span><span class="sxs-lookup"><span data-stu-id="1a9ce-222">Type</span></span>

*   [<span data-ttu-id="1a9ce-223">受信者</span><span class="sxs-lookup"><span data-stu-id="1a9ce-223">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="1a9ce-224">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-224">Requirements</span></span>

|<span data-ttu-id="1a9ce-225">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-225">Requirement</span></span>| <span data-ttu-id="1a9ce-226">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-227">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-228">1.1</span><span class="sxs-lookup"><span data-stu-id="1a9ce-228">1.1</span></span>|
|[<span data-ttu-id="1a9ce-229">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-230">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-231">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-232">作成</span><span class="sxs-lookup"><span data-stu-id="1a9ce-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1a9ce-233">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-233">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="1a9ce-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="1a9ce-235">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="1a9ce-236">タイプ</span><span class="sxs-lookup"><span data-stu-id="1a9ce-236">Type</span></span>

*   [<span data-ttu-id="1a9ce-237">Body</span><span class="sxs-lookup"><span data-stu-id="1a9ce-237">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="1a9ce-238">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-238">Requirements</span></span>

|<span data-ttu-id="1a9ce-239">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-239">Requirement</span></span>| <span data-ttu-id="1a9ce-240">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-241">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-242">1.1</span><span class="sxs-lookup"><span data-stu-id="1a9ce-242">1.1</span></span>|
|[<span data-ttu-id="1a9ce-243">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-244">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-245">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-246">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1a9ce-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1a9ce-247">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-247">Example</span></span>

<span data-ttu-id="1a9ce-248">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-248">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="1a9ce-249">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-249">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="1a9ce-250">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-250">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="1a9ce-251">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="1a9ce-252">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1a9ce-253">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-253">Read mode</span></span>

<span data-ttu-id="1a9ce-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="1a9ce-256">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-256">Compose mode</span></span>

<span data-ttu-id="1a9ce-257">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-257">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1a9ce-258">タイプ</span><span class="sxs-lookup"><span data-stu-id="1a9ce-258">Type</span></span>

*   <span data-ttu-id="1a9ce-259">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-259">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-260">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-260">Requirements</span></span>

|<span data-ttu-id="1a9ce-261">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-261">Requirement</span></span>| <span data-ttu-id="1a9ce-262">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-263">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-264">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-264">1.0</span></span>|
|[<span data-ttu-id="1a9ce-265">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-266">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-267">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-268">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1a9ce-268">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="1a9ce-269">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-269">(nullable) conversationId :String</span></span>

<span data-ttu-id="1a9ce-270">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="1a9ce-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="1a9ce-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="1a9ce-275">Type</span><span class="sxs-lookup"><span data-stu-id="1a9ce-275">Type</span></span>

*   <span data-ttu-id="1a9ce-276">String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-277">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-277">Requirements</span></span>

|<span data-ttu-id="1a9ce-278">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-278">Requirement</span></span>| <span data-ttu-id="1a9ce-279">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-280">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-281">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-281">1.0</span></span>|
|[<span data-ttu-id="1a9ce-282">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-283">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-284">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-285">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1a9ce-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1a9ce-286">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-286">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="1a9ce-287">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="1a9ce-287">dateTimeCreated :Date</span></span>

<span data-ttu-id="1a9ce-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1a9ce-290">型</span><span class="sxs-lookup"><span data-stu-id="1a9ce-290">Type</span></span>

*   <span data-ttu-id="1a9ce-291">日付</span><span class="sxs-lookup"><span data-stu-id="1a9ce-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-292">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-292">Requirements</span></span>

|<span data-ttu-id="1a9ce-293">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-293">Requirement</span></span>| <span data-ttu-id="1a9ce-294">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-295">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-296">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-296">1.0</span></span>|
|[<span data-ttu-id="1a9ce-297">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-298">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-299">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-300">読み取り</span><span class="sxs-lookup"><span data-stu-id="1a9ce-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1a9ce-301">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-301">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="1a9ce-302">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="1a9ce-302">dateTimeModified :Date</span></span>

<span data-ttu-id="1a9ce-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1a9ce-305">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-305">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="1a9ce-306">タイプ</span><span class="sxs-lookup"><span data-stu-id="1a9ce-306">Type</span></span>

*   <span data-ttu-id="1a9ce-307">日付</span><span class="sxs-lookup"><span data-stu-id="1a9ce-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-308">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-308">Requirements</span></span>

|<span data-ttu-id="1a9ce-309">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-309">Requirement</span></span>| <span data-ttu-id="1a9ce-310">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-312">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-312">1.0</span></span>|
|[<span data-ttu-id="1a9ce-313">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-314">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-315">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-316">読み取り</span><span class="sxs-lookup"><span data-stu-id="1a9ce-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1a9ce-317">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-317">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="1a9ce-318">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-318">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="1a9ce-319">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="1a9ce-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1a9ce-322">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-322">Read mode</span></span>

<span data-ttu-id="1a9ce-323">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-323">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="1a9ce-324">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-324">Compose mode</span></span>

<span data-ttu-id="1a9ce-325">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="1a9ce-326">[`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-326">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="1a9ce-327">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="1a9ce-328">タイプ</span><span class="sxs-lookup"><span data-stu-id="1a9ce-328">Type</span></span>

*   <span data-ttu-id="1a9ce-329">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-329">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-330">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-330">Requirements</span></span>

|<span data-ttu-id="1a9ce-331">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-331">Requirement</span></span>| <span data-ttu-id="1a9ce-332">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-333">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-334">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-334">1.0</span></span>|
|[<span data-ttu-id="1a9ce-335">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-336">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-337">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-338">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1a9ce-338">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="1a9ce-339">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-339">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="1a9ce-p112">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="1a9ce-p113">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1a9ce-344">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1a9ce-345">型</span><span class="sxs-lookup"><span data-stu-id="1a9ce-345">Type</span></span>

*   [<span data-ttu-id="1a9ce-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1a9ce-346">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="example"></a><span data-ttu-id="1a9ce-347">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-347">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="1a9ce-348">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-348">Requirements</span></span>

|<span data-ttu-id="1a9ce-349">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-349">Requirement</span></span>| <span data-ttu-id="1a9ce-350">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-351">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-352">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-352">1.0</span></span>|
|[<span data-ttu-id="1a9ce-353">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-354">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-355">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-356">読み取り</span><span class="sxs-lookup"><span data-stu-id="1a9ce-356">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="1a9ce-357">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-357">internetMessageId :String</span></span>

<span data-ttu-id="1a9ce-p114">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1a9ce-360">Type</span><span class="sxs-lookup"><span data-stu-id="1a9ce-360">Type</span></span>

*   <span data-ttu-id="1a9ce-361">String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-362">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-362">Requirements</span></span>

|<span data-ttu-id="1a9ce-363">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-363">Requirement</span></span>| <span data-ttu-id="1a9ce-364">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-365">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-366">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-366">1.0</span></span>|
|[<span data-ttu-id="1a9ce-367">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-368">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-369">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-370">読み取り</span><span class="sxs-lookup"><span data-stu-id="1a9ce-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1a9ce-371">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-371">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="1a9ce-372">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-372">itemClass :String</span></span>

<span data-ttu-id="1a9ce-p115">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="1a9ce-p116">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="1a9ce-377">型</span><span class="sxs-lookup"><span data-stu-id="1a9ce-377">Type</span></span> | <span data-ttu-id="1a9ce-378">説明</span><span class="sxs-lookup"><span data-stu-id="1a9ce-378">Description</span></span> | <span data-ttu-id="1a9ce-379">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="1a9ce-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="1a9ce-380">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="1a9ce-380">Appointment items</span></span> | <span data-ttu-id="1a9ce-381">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="1a9ce-382">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="1a9ce-382">Message items</span></span> | <span data-ttu-id="1a9ce-383">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="1a9ce-384">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="1a9ce-385">Type</span><span class="sxs-lookup"><span data-stu-id="1a9ce-385">Type</span></span>

*   <span data-ttu-id="1a9ce-386">String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-387">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-387">Requirements</span></span>

|<span data-ttu-id="1a9ce-388">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-388">Requirement</span></span>| <span data-ttu-id="1a9ce-389">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-390">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-391">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-391">1.0</span></span>|
|[<span data-ttu-id="1a9ce-392">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-393">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-394">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-395">読み取り</span><span class="sxs-lookup"><span data-stu-id="1a9ce-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1a9ce-396">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-396">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="1a9ce-397">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-397">(nullable) itemId :String</span></span>

<span data-ttu-id="1a9ce-p117">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1a9ce-400">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-400">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="1a9ce-401">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="1a9ce-402">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-402">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="1a9ce-403">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="1a9ce-p119">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="1a9ce-406">型</span><span class="sxs-lookup"><span data-stu-id="1a9ce-406">Type</span></span>

*   <span data-ttu-id="1a9ce-407">String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-408">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-408">Requirements</span></span>

|<span data-ttu-id="1a9ce-409">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-409">Requirement</span></span>| <span data-ttu-id="1a9ce-410">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-411">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-412">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-412">1.0</span></span>|
|[<span data-ttu-id="1a9ce-413">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-414">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-415">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-416">読み取り</span><span class="sxs-lookup"><span data-stu-id="1a9ce-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1a9ce-417">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-417">Example</span></span>

<span data-ttu-id="1a9ce-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="1a9ce-420">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-420">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="1a9ce-421">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-421">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="1a9ce-422">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-422">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="1a9ce-423">タイプ</span><span class="sxs-lookup"><span data-stu-id="1a9ce-423">Type</span></span>

*   [<span data-ttu-id="1a9ce-424">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="1a9ce-424">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="1a9ce-425">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-425">Requirements</span></span>

|<span data-ttu-id="1a9ce-426">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-426">Requirement</span></span>| <span data-ttu-id="1a9ce-427">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-428">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-429">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-429">1.0</span></span>|
|[<span data-ttu-id="1a9ce-430">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-430">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-431">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-432">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-432">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-433">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1a9ce-433">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1a9ce-434">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-434">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="1a9ce-435">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-435">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="1a9ce-436">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-436">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1a9ce-437">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-437">Read mode</span></span>

<span data-ttu-id="1a9ce-438">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-438">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="1a9ce-439">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-439">Compose mode</span></span>

<span data-ttu-id="1a9ce-440">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-440">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1a9ce-441">タイプ</span><span class="sxs-lookup"><span data-stu-id="1a9ce-441">Type</span></span>

*   <span data-ttu-id="1a9ce-442">String | [Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-442">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-443">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-443">Requirements</span></span>

|<span data-ttu-id="1a9ce-444">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-444">Requirement</span></span>| <span data-ttu-id="1a9ce-445">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-446">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-447">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-447">1.0</span></span>|
|[<span data-ttu-id="1a9ce-448">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-448">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-449">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-450">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-450">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-451">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1a9ce-451">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="1a9ce-452">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-452">normalizedSubject :String</span></span>

<span data-ttu-id="1a9ce-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="1a9ce-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="1a9ce-457">Type</span><span class="sxs-lookup"><span data-stu-id="1a9ce-457">Type</span></span>

*   <span data-ttu-id="1a9ce-458">String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-458">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-459">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-459">Requirements</span></span>

|<span data-ttu-id="1a9ce-460">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-460">Requirement</span></span>| <span data-ttu-id="1a9ce-461">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-462">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-463">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-463">1.0</span></span>|
|[<span data-ttu-id="1a9ce-464">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-464">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-465">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-466">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-466">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-467">読み取り</span><span class="sxs-lookup"><span data-stu-id="1a9ce-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1a9ce-468">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-468">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="1a9ce-469">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-469">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="1a9ce-470">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-470">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="1a9ce-471">型</span><span class="sxs-lookup"><span data-stu-id="1a9ce-471">Type</span></span>

*   [<span data-ttu-id="1a9ce-472">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="1a9ce-472">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="1a9ce-473">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-473">Requirements</span></span>

|<span data-ttu-id="1a9ce-474">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-474">Requirement</span></span>| <span data-ttu-id="1a9ce-475">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-475">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-476">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-476">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-477">1.3</span><span class="sxs-lookup"><span data-stu-id="1a9ce-477">1.3</span></span>|
|[<span data-ttu-id="1a9ce-478">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-478">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-479">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-479">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-480">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-480">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-481">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1a9ce-481">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1a9ce-482">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-482">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="1a9ce-483">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-483">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="1a9ce-484">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-484">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="1a9ce-485">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-485">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1a9ce-486">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-486">Read mode</span></span>

<span data-ttu-id="1a9ce-487">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-487">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="1a9ce-488">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-488">Compose mode</span></span>

<span data-ttu-id="1a9ce-489">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-489">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1a9ce-490">タイプ</span><span class="sxs-lookup"><span data-stu-id="1a9ce-490">Type</span></span>

*   <span data-ttu-id="1a9ce-491">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-491">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-492">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-492">Requirements</span></span>

|<span data-ttu-id="1a9ce-493">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-493">Requirement</span></span>| <span data-ttu-id="1a9ce-494">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-494">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-495">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-495">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-496">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-496">1.0</span></span>|
|[<span data-ttu-id="1a9ce-497">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-497">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-498">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-498">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-499">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-499">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-500">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1a9ce-500">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="1a9ce-501">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-501">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="1a9ce-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1a9ce-504">タイプ</span><span class="sxs-lookup"><span data-stu-id="1a9ce-504">Type</span></span>

*   [<span data-ttu-id="1a9ce-505">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1a9ce-505">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="1a9ce-506">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-506">Requirements</span></span>

|<span data-ttu-id="1a9ce-507">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-507">Requirement</span></span>| <span data-ttu-id="1a9ce-508">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-509">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-510">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-510">1.0</span></span>|
|[<span data-ttu-id="1a9ce-511">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-512">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-513">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-514">読み取り</span><span class="sxs-lookup"><span data-stu-id="1a9ce-514">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1a9ce-515">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-515">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="1a9ce-516">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-516">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="1a9ce-517">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-517">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="1a9ce-518">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-518">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1a9ce-519">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-519">Read mode</span></span>

<span data-ttu-id="1a9ce-520">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-520">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="1a9ce-521">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-521">Compose mode</span></span>

<span data-ttu-id="1a9ce-522">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-522">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="1a9ce-523">型</span><span class="sxs-lookup"><span data-stu-id="1a9ce-523">Type</span></span>

*   <span data-ttu-id="1a9ce-524">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-524">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-525">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-525">Requirements</span></span>

|<span data-ttu-id="1a9ce-526">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-526">Requirement</span></span>| <span data-ttu-id="1a9ce-527">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-527">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-528">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-528">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-529">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-529">1.0</span></span>|
|[<span data-ttu-id="1a9ce-530">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-530">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-531">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-532">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-532">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-533">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1a9ce-533">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="1a9ce-534">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-534">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="1a9ce-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="1a9ce-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1a9ce-539">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-539">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1a9ce-540">タイプ</span><span class="sxs-lookup"><span data-stu-id="1a9ce-540">Type</span></span>

*   [<span data-ttu-id="1a9ce-541">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1a9ce-541">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="1a9ce-542">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-542">Requirements</span></span>

|<span data-ttu-id="1a9ce-543">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-543">Requirement</span></span>| <span data-ttu-id="1a9ce-544">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-545">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-546">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-546">1.0</span></span>|
|[<span data-ttu-id="1a9ce-547">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-548">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-549">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-550">読み取り</span><span class="sxs-lookup"><span data-stu-id="1a9ce-550">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1a9ce-551">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-551">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="1a9ce-552">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-552">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="1a9ce-553">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-553">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="1a9ce-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1a9ce-556">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-556">Read mode</span></span>

<span data-ttu-id="1a9ce-557">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-557">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="1a9ce-558">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-558">Compose mode</span></span>

<span data-ttu-id="1a9ce-559">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-559">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="1a9ce-560">[`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-560">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="1a9ce-561">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-561">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="1a9ce-562">タイプ</span><span class="sxs-lookup"><span data-stu-id="1a9ce-562">Type</span></span>

*   <span data-ttu-id="1a9ce-563">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-563">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-564">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-564">Requirements</span></span>

|<span data-ttu-id="1a9ce-565">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-565">Requirement</span></span>| <span data-ttu-id="1a9ce-566">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-567">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-568">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-568">1.0</span></span>|
|[<span data-ttu-id="1a9ce-569">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-570">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-571">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-572">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1a9ce-572">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="1a9ce-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="1a9ce-574">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-574">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="1a9ce-575">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-575">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1a9ce-576">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-576">Read mode</span></span>

<span data-ttu-id="1a9ce-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="1a9ce-579">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-579">Compose mode</span></span>

<span data-ttu-id="1a9ce-580">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-580">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="1a9ce-581">型</span><span class="sxs-lookup"><span data-stu-id="1a9ce-581">Type</span></span>

*   <span data-ttu-id="1a9ce-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-583">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-583">Requirements</span></span>

|<span data-ttu-id="1a9ce-584">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-584">Requirement</span></span>| <span data-ttu-id="1a9ce-585">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-585">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-586">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-586">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-587">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-587">1.0</span></span>|
|[<span data-ttu-id="1a9ce-588">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-588">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-589">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-589">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-590">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-590">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-591">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1a9ce-591">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="1a9ce-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="1a9ce-593">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-593">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="1a9ce-594">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-594">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1a9ce-595">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-595">Read mode</span></span>

<span data-ttu-id="1a9ce-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="1a9ce-598">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-598">Compose mode</span></span>

<span data-ttu-id="1a9ce-599">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-599">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1a9ce-600">型</span><span class="sxs-lookup"><span data-stu-id="1a9ce-600">Type</span></span>

*   <span data-ttu-id="1a9ce-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-602">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-602">Requirements</span></span>

|<span data-ttu-id="1a9ce-603">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-603">Requirement</span></span>| <span data-ttu-id="1a9ce-604">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-605">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-606">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-606">1.0</span></span>|
|[<span data-ttu-id="1a9ce-607">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-608">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-608">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-609">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-610">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1a9ce-610">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="1a9ce-611">メソッド</span><span class="sxs-lookup"><span data-stu-id="1a9ce-611">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="1a9ce-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1a9ce-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1a9ce-613">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="1a9ce-614">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="1a9ce-615">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a9ce-616">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1a9ce-616">Parameters</span></span>

|<span data-ttu-id="1a9ce-617">名前</span><span class="sxs-lookup"><span data-stu-id="1a9ce-617">Name</span></span>| <span data-ttu-id="1a9ce-618">種類</span><span class="sxs-lookup"><span data-stu-id="1a9ce-618">Type</span></span>| <span data-ttu-id="1a9ce-619">属性</span><span class="sxs-lookup"><span data-stu-id="1a9ce-619">Attributes</span></span>| <span data-ttu-id="1a9ce-620">説明</span><span class="sxs-lookup"><span data-stu-id="1a9ce-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="1a9ce-621">文字列</span><span class="sxs-lookup"><span data-stu-id="1a9ce-621">String</span></span>||<span data-ttu-id="1a9ce-p132">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="1a9ce-624">String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-624">String</span></span>||<span data-ttu-id="1a9ce-p133">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="1a9ce-627">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1a9ce-627">Object</span></span>| <span data-ttu-id="1a9ce-628">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-628">&lt;optional&gt;</span></span>|<span data-ttu-id="1a9ce-629">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="1a9ce-630">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1a9ce-630">Object</span></span> | <span data-ttu-id="1a9ce-631">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-631">&lt;optional&gt;</span></span> | <span data-ttu-id="1a9ce-632">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="1a9ce-633">Boolean</span><span class="sxs-lookup"><span data-stu-id="1a9ce-633">Boolean</span></span> | <span data-ttu-id="1a9ce-634">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-634">&lt;optional&gt;</span></span> | <span data-ttu-id="1a9ce-635">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="1a9ce-636">function</span><span class="sxs-lookup"><span data-stu-id="1a9ce-636">function</span></span>| <span data-ttu-id="1a9ce-637">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-637">&lt;optional&gt;</span></span>|<span data-ttu-id="1a9ce-638">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1a9ce-639">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1a9ce-640">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1a9ce-641">エラー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-641">Errors</span></span>

| <span data-ttu-id="1a9ce-642">エラー コード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-642">Error code</span></span> | <span data-ttu-id="1a9ce-643">説明</span><span class="sxs-lookup"><span data-stu-id="1a9ce-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="1a9ce-644">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="1a9ce-645">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="1a9ce-646">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1a9ce-647">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-647">Requirements</span></span>

|<span data-ttu-id="1a9ce-648">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-648">Requirement</span></span>| <span data-ttu-id="1a9ce-649">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-650">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-651">1.1</span><span class="sxs-lookup"><span data-stu-id="1a9ce-651">1.1</span></span>|
|[<span data-ttu-id="1a9ce-652">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-652">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="1a9ce-654">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-654">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-655">作成</span><span class="sxs-lookup"><span data-stu-id="1a9ce-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="1a9ce-656">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-656">Examples</span></span>

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

<span data-ttu-id="1a9ce-657">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="1a9ce-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1a9ce-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1a9ce-659">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="1a9ce-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="1a9ce-663">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="1a9ce-664">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-664">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a9ce-665">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1a9ce-665">Parameters</span></span>

|<span data-ttu-id="1a9ce-666">名前</span><span class="sxs-lookup"><span data-stu-id="1a9ce-666">Name</span></span>| <span data-ttu-id="1a9ce-667">型</span><span class="sxs-lookup"><span data-stu-id="1a9ce-667">Type</span></span>| <span data-ttu-id="1a9ce-668">属性</span><span class="sxs-lookup"><span data-stu-id="1a9ce-668">Attributes</span></span>| <span data-ttu-id="1a9ce-669">説明</span><span class="sxs-lookup"><span data-stu-id="1a9ce-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="1a9ce-670">String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-670">String</span></span>||<span data-ttu-id="1a9ce-p135">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="1a9ce-673">String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-673">String</span></span>||<span data-ttu-id="1a9ce-674">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-674">The subject of the item to be attached.</span></span> <span data-ttu-id="1a9ce-675">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-675">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="1a9ce-676">Object</span><span class="sxs-lookup"><span data-stu-id="1a9ce-676">Object</span></span>| <span data-ttu-id="1a9ce-677">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-677">&lt;optional&gt;</span></span>|<span data-ttu-id="1a9ce-678">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1a9ce-679">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1a9ce-679">Object</span></span>| <span data-ttu-id="1a9ce-680">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-680">&lt;optional&gt;</span></span>|<span data-ttu-id="1a9ce-681">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1a9ce-682">関数</span><span class="sxs-lookup"><span data-stu-id="1a9ce-682">function</span></span>| <span data-ttu-id="1a9ce-683">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-683">&lt;optional&gt;</span></span>|<span data-ttu-id="1a9ce-684">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1a9ce-685">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1a9ce-686">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1a9ce-687">エラー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-687">Errors</span></span>

| <span data-ttu-id="1a9ce-688">エラー コード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-688">Error code</span></span> | <span data-ttu-id="1a9ce-689">説明</span><span class="sxs-lookup"><span data-stu-id="1a9ce-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="1a9ce-690">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1a9ce-691">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-691">Requirements</span></span>

|<span data-ttu-id="1a9ce-692">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-692">Requirement</span></span>| <span data-ttu-id="1a9ce-693">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-694">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-695">1.1</span><span class="sxs-lookup"><span data-stu-id="1a9ce-695">1.1</span></span>|
|[<span data-ttu-id="1a9ce-696">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-696">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="1a9ce-698">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-698">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-699">作成</span><span class="sxs-lookup"><span data-stu-id="1a9ce-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1a9ce-700">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-700">Example</span></span>

<span data-ttu-id="1a9ce-701">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="1a9ce-702">close()</span><span class="sxs-lookup"><span data-stu-id="1a9ce-702">close()</span></span>

<span data-ttu-id="1a9ce-703">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="1a9ce-p137">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="1a9ce-706">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="1a9ce-707">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-708">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-708">Requirements</span></span>

|<span data-ttu-id="1a9ce-709">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-709">Requirement</span></span>| <span data-ttu-id="1a9ce-710">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-711">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-712">1.3</span><span class="sxs-lookup"><span data-stu-id="1a9ce-712">1.3</span></span>|
|[<span data-ttu-id="1a9ce-713">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-713">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-714">制限あり</span><span class="sxs-lookup"><span data-stu-id="1a9ce-714">Restricted</span></span>|
|[<span data-ttu-id="1a9ce-715">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-715">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-716">新規作成</span><span class="sxs-lookup"><span data-stu-id="1a9ce-716">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="1a9ce-717">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="1a9ce-717">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="1a9ce-718">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1a9ce-719">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-719">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1a9ce-720">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-720">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1a9ce-721">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="1a9ce-p138">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a9ce-725">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1a9ce-725">Parameters</span></span>

| <span data-ttu-id="1a9ce-726">名前</span><span class="sxs-lookup"><span data-stu-id="1a9ce-726">Name</span></span> | <span data-ttu-id="1a9ce-727">型</span><span class="sxs-lookup"><span data-stu-id="1a9ce-727">Type</span></span> | <span data-ttu-id="1a9ce-728">属性</span><span class="sxs-lookup"><span data-stu-id="1a9ce-728">Attributes</span></span> | <span data-ttu-id="1a9ce-729">説明</span><span class="sxs-lookup"><span data-stu-id="1a9ce-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="1a9ce-730">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="1a9ce-730">String &#124; Object</span></span>| |<span data-ttu-id="1a9ce-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1a9ce-733">**または**</span><span class="sxs-lookup"><span data-stu-id="1a9ce-733">**OR**</span></span><br/><span data-ttu-id="1a9ce-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="1a9ce-736">String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-736">String</span></span> | <span data-ttu-id="1a9ce-737">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-737">&lt;optional&gt;</span></span> | <span data-ttu-id="1a9ce-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="1a9ce-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="1a9ce-741">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-741">&lt;optional&gt;</span></span> | <span data-ttu-id="1a9ce-742">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="1a9ce-743">文字列</span><span class="sxs-lookup"><span data-stu-id="1a9ce-743">String</span></span> | | <span data-ttu-id="1a9ce-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="1a9ce-746">String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-746">String</span></span> | | <span data-ttu-id="1a9ce-747">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="1a9ce-748">文字列</span><span class="sxs-lookup"><span data-stu-id="1a9ce-748">String</span></span> | | <span data-ttu-id="1a9ce-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="1a9ce-751">ブール値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-751">Boolean</span></span> | | <span data-ttu-id="1a9ce-p144">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="1a9ce-754">String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-754">String</span></span> | | <span data-ttu-id="1a9ce-p145">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="1a9ce-758">function</span><span class="sxs-lookup"><span data-stu-id="1a9ce-758">function</span></span> | <span data-ttu-id="1a9ce-759">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-759">&lt;optional&gt;</span></span> | <span data-ttu-id="1a9ce-760">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1a9ce-761">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-761">Requirements</span></span>

|<span data-ttu-id="1a9ce-762">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-762">Requirement</span></span>| <span data-ttu-id="1a9ce-763">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-764">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-765">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-765">1.0</span></span>|
|[<span data-ttu-id="1a9ce-766">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-766">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-767">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-768">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-768">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-769">読み取り</span><span class="sxs-lookup"><span data-stu-id="1a9ce-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1a9ce-770">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-770">Examples</span></span>

<span data-ttu-id="1a9ce-771">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="1a9ce-772">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-772">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="1a9ce-773">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-773">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1a9ce-774">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-774">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="1a9ce-775">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-775">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="1a9ce-776">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="1a9ce-777">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="1a9ce-777">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="1a9ce-778">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1a9ce-779">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-779">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1a9ce-780">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-780">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1a9ce-781">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="1a9ce-p146">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a9ce-785">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1a9ce-785">Parameters</span></span>

| <span data-ttu-id="1a9ce-786">名前</span><span class="sxs-lookup"><span data-stu-id="1a9ce-786">Name</span></span> | <span data-ttu-id="1a9ce-787">型</span><span class="sxs-lookup"><span data-stu-id="1a9ce-787">Type</span></span> | <span data-ttu-id="1a9ce-788">属性</span><span class="sxs-lookup"><span data-stu-id="1a9ce-788">Attributes</span></span> | <span data-ttu-id="1a9ce-789">説明</span><span class="sxs-lookup"><span data-stu-id="1a9ce-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="1a9ce-790">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="1a9ce-790">String &#124; Object</span></span>| | <span data-ttu-id="1a9ce-p147">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1a9ce-793">**または**</span><span class="sxs-lookup"><span data-stu-id="1a9ce-793">**OR**</span></span><br/><span data-ttu-id="1a9ce-p148">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="1a9ce-796">String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-796">String</span></span> | <span data-ttu-id="1a9ce-797">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-797">&lt;optional&gt;</span></span> | <span data-ttu-id="1a9ce-p149">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="1a9ce-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="1a9ce-801">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-801">&lt;optional&gt;</span></span> | <span data-ttu-id="1a9ce-802">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="1a9ce-803">文字列</span><span class="sxs-lookup"><span data-stu-id="1a9ce-803">String</span></span> | | <span data-ttu-id="1a9ce-p150">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="1a9ce-806">String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-806">String</span></span> | | <span data-ttu-id="1a9ce-807">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="1a9ce-808">文字列</span><span class="sxs-lookup"><span data-stu-id="1a9ce-808">String</span></span> | | <span data-ttu-id="1a9ce-p151">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="1a9ce-811">ブール値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-811">Boolean</span></span> | | <span data-ttu-id="1a9ce-p152">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="1a9ce-814">String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-814">String</span></span> | | <span data-ttu-id="1a9ce-p153">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="1a9ce-818">function</span><span class="sxs-lookup"><span data-stu-id="1a9ce-818">function</span></span> | <span data-ttu-id="1a9ce-819">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-819">&lt;optional&gt;</span></span> | <span data-ttu-id="1a9ce-820">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1a9ce-821">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-821">Requirements</span></span>

|<span data-ttu-id="1a9ce-822">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-822">Requirement</span></span>| <span data-ttu-id="1a9ce-823">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-824">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-824">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-825">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-825">1.0</span></span>|
|[<span data-ttu-id="1a9ce-826">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-826">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-827">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-828">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-828">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-829">読み取り</span><span class="sxs-lookup"><span data-stu-id="1a9ce-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1a9ce-830">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-830">Examples</span></span>

<span data-ttu-id="1a9ce-831">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="1a9ce-832">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-832">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="1a9ce-833">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-833">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1a9ce-834">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-834">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="1a9ce-835">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-835">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="1a9ce-836">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="1a9ce-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="1a9ce-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="1a9ce-838">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-838">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="1a9ce-839">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-839">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-840">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-840">Requirements</span></span>

|<span data-ttu-id="1a9ce-841">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-841">Requirement</span></span>| <span data-ttu-id="1a9ce-842">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-843">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-844">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-844">1.0</span></span>|
|[<span data-ttu-id="1a9ce-845">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-845">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-846">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-847">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-847">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-848">読み取り</span><span class="sxs-lookup"><span data-stu-id="1a9ce-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1a9ce-849">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1a9ce-849">Returns:</span></span>

<span data-ttu-id="1a9ce-850">型:[Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-850">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="1a9ce-851">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-851">Example</span></span>

<span data-ttu-id="1a9ce-852">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-852">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="1a9ce-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="1a9ce-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="1a9ce-854">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-854">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="1a9ce-855">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-855">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a9ce-856">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1a9ce-856">Parameters</span></span>

|<span data-ttu-id="1a9ce-857">名前</span><span class="sxs-lookup"><span data-stu-id="1a9ce-857">Name</span></span>| <span data-ttu-id="1a9ce-858">型</span><span class="sxs-lookup"><span data-stu-id="1a9ce-858">Type</span></span>| <span data-ttu-id="1a9ce-859">説明</span><span class="sxs-lookup"><span data-stu-id="1a9ce-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="1a9ce-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="1a9ce-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="1a9ce-861">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1a9ce-862">Requirements</span><span class="sxs-lookup"><span data-stu-id="1a9ce-862">Requirements</span></span>

|<span data-ttu-id="1a9ce-863">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-863">Requirement</span></span>| <span data-ttu-id="1a9ce-864">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-865">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-866">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-866">1.0</span></span>|
|[<span data-ttu-id="1a9ce-867">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-867">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-868">制限あり</span><span class="sxs-lookup"><span data-stu-id="1a9ce-868">Restricted</span></span>|
|[<span data-ttu-id="1a9ce-869">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-869">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-870">読み取り</span><span class="sxs-lookup"><span data-stu-id="1a9ce-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1a9ce-871">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1a9ce-871">Returns:</span></span>

<span data-ttu-id="1a9ce-872">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="1a9ce-873">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-873">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="1a9ce-874">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="1a9ce-875">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="1a9ce-876">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-876">Value of `entityType`</span></span> | <span data-ttu-id="1a9ce-877">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="1a9ce-877">Type of objects in returned array</span></span> | <span data-ttu-id="1a9ce-878">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="1a9ce-879">文字列</span><span class="sxs-lookup"><span data-stu-id="1a9ce-879">String</span></span> | <span data-ttu-id="1a9ce-880">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="1a9ce-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="1a9ce-881">連絡先</span><span class="sxs-lookup"><span data-stu-id="1a9ce-881">Contact</span></span> | <span data-ttu-id="1a9ce-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1a9ce-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="1a9ce-883">文字列</span><span class="sxs-lookup"><span data-stu-id="1a9ce-883">String</span></span> | <span data-ttu-id="1a9ce-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1a9ce-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="1a9ce-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="1a9ce-885">MeetingSuggestion</span></span> | <span data-ttu-id="1a9ce-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1a9ce-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="1a9ce-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="1a9ce-887">PhoneNumber</span></span> | <span data-ttu-id="1a9ce-888">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="1a9ce-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="1a9ce-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="1a9ce-889">TaskSuggestion</span></span> | <span data-ttu-id="1a9ce-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1a9ce-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="1a9ce-891">文字列</span><span class="sxs-lookup"><span data-stu-id="1a9ce-891">String</span></span> | <span data-ttu-id="1a9ce-892">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="1a9ce-892">**Restricted**</span></span> |

<span data-ttu-id="1a9ce-893">型:Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="1a9ce-893">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="1a9ce-894">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-894">Example</span></span>

<span data-ttu-id="1a9ce-895">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-895">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="1a9ce-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="1a9ce-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="1a9ce-897">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1a9ce-898">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-898">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1a9ce-899">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a9ce-900">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1a9ce-900">Parameters</span></span>

|<span data-ttu-id="1a9ce-901">名前</span><span class="sxs-lookup"><span data-stu-id="1a9ce-901">Name</span></span>| <span data-ttu-id="1a9ce-902">型</span><span class="sxs-lookup"><span data-stu-id="1a9ce-902">Type</span></span>| <span data-ttu-id="1a9ce-903">説明</span><span class="sxs-lookup"><span data-stu-id="1a9ce-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="1a9ce-904">String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-904">String</span></span>|<span data-ttu-id="1a9ce-905">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1a9ce-906">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-906">Requirements</span></span>

|<span data-ttu-id="1a9ce-907">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-907">Requirement</span></span>| <span data-ttu-id="1a9ce-908">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-909">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-909">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-910">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-910">1.0</span></span>|
|[<span data-ttu-id="1a9ce-911">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-911">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-912">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-913">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-913">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-914">読み取り</span><span class="sxs-lookup"><span data-stu-id="1a9ce-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1a9ce-915">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1a9ce-915">Returns:</span></span>

<span data-ttu-id="1a9ce-p155">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="1a9ce-918">型:Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="1a9ce-918">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="1a9ce-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="1a9ce-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="1a9ce-920">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1a9ce-921">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-921">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1a9ce-p156">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="1a9ce-925">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="1a9ce-926">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="1a9ce-p157">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-930">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-930">Requirements</span></span>

|<span data-ttu-id="1a9ce-931">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-931">Requirement</span></span>| <span data-ttu-id="1a9ce-932">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-933">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-934">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-934">1.0</span></span>|
|[<span data-ttu-id="1a9ce-935">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-936">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-937">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-938">読み取り</span><span class="sxs-lookup"><span data-stu-id="1a9ce-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1a9ce-939">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1a9ce-939">Returns:</span></span>

<span data-ttu-id="1a9ce-p158">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="1a9ce-942">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="1a9ce-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1a9ce-943">Object</span><span class="sxs-lookup"><span data-stu-id="1a9ce-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1a9ce-944">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-944">Example</span></span>

<span data-ttu-id="1a9ce-945">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="1a9ce-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="1a9ce-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="1a9ce-947">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1a9ce-948">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-948">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1a9ce-949">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="1a9ce-p159">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a9ce-952">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1a9ce-952">Parameters</span></span>

|<span data-ttu-id="1a9ce-953">名前</span><span class="sxs-lookup"><span data-stu-id="1a9ce-953">Name</span></span>| <span data-ttu-id="1a9ce-954">型</span><span class="sxs-lookup"><span data-stu-id="1a9ce-954">Type</span></span>| <span data-ttu-id="1a9ce-955">説明</span><span class="sxs-lookup"><span data-stu-id="1a9ce-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="1a9ce-956">String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-956">String</span></span>|<span data-ttu-id="1a9ce-957">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1a9ce-958">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-958">Requirements</span></span>

|<span data-ttu-id="1a9ce-959">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-959">Requirement</span></span>| <span data-ttu-id="1a9ce-960">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-961">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-961">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-962">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-962">1.0</span></span>|
|[<span data-ttu-id="1a9ce-963">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-963">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-964">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-965">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-965">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-966">読み取り</span><span class="sxs-lookup"><span data-stu-id="1a9ce-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1a9ce-967">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1a9ce-967">Returns:</span></span>

<span data-ttu-id="1a9ce-968">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="1a9ce-969">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="1a9ce-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1a9ce-970">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="1a9ce-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1a9ce-971">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-971">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="1a9ce-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="1a9ce-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="1a9ce-973">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="1a9ce-p160">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a9ce-976">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1a9ce-976">Parameters</span></span>

|<span data-ttu-id="1a9ce-977">名前</span><span class="sxs-lookup"><span data-stu-id="1a9ce-977">Name</span></span>| <span data-ttu-id="1a9ce-978">型</span><span class="sxs-lookup"><span data-stu-id="1a9ce-978">Type</span></span>| <span data-ttu-id="1a9ce-979">属性</span><span class="sxs-lookup"><span data-stu-id="1a9ce-979">Attributes</span></span>| <span data-ttu-id="1a9ce-980">説明</span><span class="sxs-lookup"><span data-stu-id="1a9ce-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="1a9ce-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1a9ce-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="1a9ce-p161">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="1a9ce-985">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1a9ce-985">Object</span></span>| <span data-ttu-id="1a9ce-986">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-986">&lt;optional&gt;</span></span>|<span data-ttu-id="1a9ce-987">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1a9ce-988">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1a9ce-988">Object</span></span>| <span data-ttu-id="1a9ce-989">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-989">&lt;optional&gt;</span></span>|<span data-ttu-id="1a9ce-990">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1a9ce-991">function</span><span class="sxs-lookup"><span data-stu-id="1a9ce-991">function</span></span>||<span data-ttu-id="1a9ce-992">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1a9ce-993">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="1a9ce-994">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-994">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1a9ce-995">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-995">Requirements</span></span>

|<span data-ttu-id="1a9ce-996">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-996">Requirement</span></span>| <span data-ttu-id="1a9ce-997">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-998">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-998">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-999">1.2</span><span class="sxs-lookup"><span data-stu-id="1a9ce-999">1.2</span></span>|
|[<span data-ttu-id="1a9ce-1000">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1000">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="1a9ce-1002">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1002">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-1003">作成</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="1a9ce-1004">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1004">Returns:</span></span>

<span data-ttu-id="1a9ce-1005">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="1a9ce-1006">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1a9ce-1007">String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1a9ce-1008">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1008">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="1a9ce-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="1a9ce-1010">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1010">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="1a9ce-1011">強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1011">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="1a9ce-1012">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1012">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-1013">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1013">Requirements</span></span>

|<span data-ttu-id="1a9ce-1014">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1014">Requirement</span></span>| <span data-ttu-id="1a9ce-1015">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-1016">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1016">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1017">1.6</span></span> |
|[<span data-ttu-id="1a9ce-1018">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1018">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1019">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-1020">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1020">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-1021">読み取り</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1a9ce-1022">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1022">Returns:</span></span>

<span data-ttu-id="1a9ce-1023">型:[Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1023">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="1a9ce-1024">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1024">Example</span></span>

<span data-ttu-id="1a9ce-1025">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="1a9ce-1026">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="1a9ce-p164">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="1a9ce-1029">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1029">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1a9ce-p165">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="1a9ce-1033">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="1a9ce-1034">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="1a9ce-p166">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a9ce-1038">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1038">Requirements</span></span>

|<span data-ttu-id="1a9ce-1039">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1039">Requirement</span></span>| <span data-ttu-id="1a9ce-1040">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-1041">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1041">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1042">1.6</span></span> |
|[<span data-ttu-id="1a9ce-1043">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1043">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1044">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-1045">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1045">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-1046">読み取り</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1a9ce-1047">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1047">Returns:</span></span>

<span data-ttu-id="1a9ce-p167">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="1a9ce-1050">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1050">Example</span></span>

<span data-ttu-id="1a9ce-1051">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="1a9ce-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="1a9ce-1053">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="1a9ce-p168">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a9ce-1057">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1057">Parameters</span></span>

|<span data-ttu-id="1a9ce-1058">名前</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1058">Name</span></span>| <span data-ttu-id="1a9ce-1059">型</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1059">Type</span></span>| <span data-ttu-id="1a9ce-1060">属性</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1060">Attributes</span></span>| <span data-ttu-id="1a9ce-1061">説明</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="1a9ce-1062">function</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1062">function</span></span>||<span data-ttu-id="1a9ce-1063">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1a9ce-1064">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="1a9ce-1065">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1065">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="1a9ce-1066">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1066">Object</span></span>| <span data-ttu-id="1a9ce-1067">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="1a9ce-1068">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1068">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="1a9ce-1069">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1a9ce-1070">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1070">Requirements</span></span>

|<span data-ttu-id="1a9ce-1071">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1071">Requirement</span></span>| <span data-ttu-id="1a9ce-1072">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-1073">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1073">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-1074">1.0</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1074">1.0</span></span>|
|[<span data-ttu-id="1a9ce-1075">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1075">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1076">ReadItem</span></span>|
|[<span data-ttu-id="1a9ce-1077">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1077">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-1078">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1078">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1a9ce-1079">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1079">Example</span></span>

<span data-ttu-id="1a9ce-p171">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="1a9ce-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="1a9ce-1084">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="1a9ce-p172">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a9ce-1089">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1089">Parameters</span></span>

|<span data-ttu-id="1a9ce-1090">名前</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1090">Name</span></span>| <span data-ttu-id="1a9ce-1091">型</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1091">Type</span></span>| <span data-ttu-id="1a9ce-1092">属性</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1092">Attributes</span></span>| <span data-ttu-id="1a9ce-1093">説明</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="1a9ce-1094">文字列</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1094">String</span></span>||<span data-ttu-id="1a9ce-1095">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1095">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="1a9ce-1096">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1096">Object</span></span>| <span data-ttu-id="1a9ce-1097">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="1a9ce-1098">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1098">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1a9ce-1099">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1099">Object</span></span>| <span data-ttu-id="1a9ce-1100">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1100">&lt;optional&gt;</span></span>|<span data-ttu-id="1a9ce-1101">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1101">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1a9ce-1102">function</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1102">function</span></span>| <span data-ttu-id="1a9ce-1103">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1103">&lt;optional&gt;</span></span>|<span data-ttu-id="1a9ce-1104">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1104">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1a9ce-1105">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1105">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1a9ce-1106">エラー</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1106">Errors</span></span>

| <span data-ttu-id="1a9ce-1107">エラー コード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1107">Error code</span></span> | <span data-ttu-id="1a9ce-1108">説明</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1108">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="1a9ce-1109">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1109">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1a9ce-1110">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1110">Requirements</span></span>

|<span data-ttu-id="1a9ce-1111">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1111">Requirement</span></span>| <span data-ttu-id="1a9ce-1112">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1112">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-1113">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1113">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-1114">1.1</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1114">1.1</span></span>|
|[<span data-ttu-id="1a9ce-1115">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1115">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-1116">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1116">ReadWriteItem</span></span>|
|[<span data-ttu-id="1a9ce-1117">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1117">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-1118">作成</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1118">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1a9ce-1119">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1119">Example</span></span>

<span data-ttu-id="1a9ce-1120">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1120">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="1a9ce-1121">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1121">saveAsync([options], callback)</span></span>

<span data-ttu-id="1a9ce-1122">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1122">Asynchronously saves an item.</span></span>

<span data-ttu-id="1a9ce-p173">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p173">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="1a9ce-1126">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1126">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="1a9ce-1127">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1127">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="1a9ce-p175">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="1a9ce-1131">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1131">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="1a9ce-1132">Mac Outlook では、新規作成モードの会議で `saveAsync` をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1132">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="1a9ce-1133">Mac Outlook では、会議で `saveAsync` を呼び出すとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1133">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="1a9ce-1134">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1134">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a9ce-1135">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1135">Parameters</span></span>

|<span data-ttu-id="1a9ce-1136">名前</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1136">Name</span></span>| <span data-ttu-id="1a9ce-1137">型</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1137">Type</span></span>| <span data-ttu-id="1a9ce-1138">属性</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1138">Attributes</span></span>| <span data-ttu-id="1a9ce-1139">説明</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1139">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="1a9ce-1140">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1140">Object</span></span>| <span data-ttu-id="1a9ce-1141">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1141">&lt;optional&gt;</span></span>|<span data-ttu-id="1a9ce-1142">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1142">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1a9ce-1143">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1143">Object</span></span>| <span data-ttu-id="1a9ce-1144">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1144">&lt;optional&gt;</span></span>|<span data-ttu-id="1a9ce-1145">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1145">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1a9ce-1146">関数</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1146">function</span></span>||<span data-ttu-id="1a9ce-1147">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1147">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1a9ce-1148">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1148">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1a9ce-1149">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1149">Requirements</span></span>

|<span data-ttu-id="1a9ce-1150">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1150">Requirement</span></span>| <span data-ttu-id="1a9ce-1151">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1151">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-1152">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1152">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-1153">1.3</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1153">1.3</span></span>|
|[<span data-ttu-id="1a9ce-1154">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1154">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-1155">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1155">ReadWriteItem</span></span>|
|[<span data-ttu-id="1a9ce-1156">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1156">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-1157">作成</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1157">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="1a9ce-1158">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1158">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="1a9ce-p177">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="1a9ce-1161">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1161">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="1a9ce-1162">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1162">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="1a9ce-p178">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1a9ce-1166">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1166">Parameters</span></span>

|<span data-ttu-id="1a9ce-1167">名前</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1167">Name</span></span>| <span data-ttu-id="1a9ce-1168">型</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1168">Type</span></span>| <span data-ttu-id="1a9ce-1169">属性</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1169">Attributes</span></span>| <span data-ttu-id="1a9ce-1170">説明</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1170">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="1a9ce-1171">String</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1171">String</span></span>||<span data-ttu-id="1a9ce-p179">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="1a9ce-1175">Object</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1175">Object</span></span>| <span data-ttu-id="1a9ce-1176">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1176">&lt;optional&gt;</span></span>|<span data-ttu-id="1a9ce-1177">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1177">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1a9ce-1178">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1178">Object</span></span>| <span data-ttu-id="1a9ce-1179">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1179">&lt;optional&gt;</span></span>|<span data-ttu-id="1a9ce-1180">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1180">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="1a9ce-1181">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1181">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="1a9ce-1182">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1182">&lt;optional&gt;</span></span>|<span data-ttu-id="1a9ce-p180">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p180">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="1a9ce-p181">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-p181">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="1a9ce-1187">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1187">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="1a9ce-1188">function</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1188">function</span></span>||<span data-ttu-id="1a9ce-1189">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1189">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1a9ce-1190">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1190">Requirements</span></span>

|<span data-ttu-id="1a9ce-1191">要件</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1191">Requirement</span></span>| <span data-ttu-id="1a9ce-1192">値</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1192">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a9ce-1193">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1193">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1a9ce-1194">1.2</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1194">1.2</span></span>|
|[<span data-ttu-id="1a9ce-1195">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1195">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1a9ce-1196">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1196">ReadWriteItem</span></span>|
|[<span data-ttu-id="1a9ce-1197">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1197">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1a9ce-1198">作成</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1198">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1a9ce-1199">例</span><span class="sxs-lookup"><span data-stu-id="1a9ce-1199">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
