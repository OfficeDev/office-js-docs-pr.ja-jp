---
title: Office.context.mailbox.item - プレビュー要件セット
description: ''
ms.date: 01/30/2019
localization_priority: Normal
ms.openlocfilehash: a660f8bafdd2587f97d704e42c47abbe6c7d533d
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/13/2019
ms.locfileid: "29982049"
---
# <a name="item"></a><span data-ttu-id="8e966-102">item</span><span class="sxs-lookup"><span data-stu-id="8e966-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="8e966-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="8e966-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="8e966-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="8e966-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-106">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-106">Requirements</span></span>

|<span data-ttu-id="8e966-107">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-107">Requirement</span></span>|<span data-ttu-id="8e966-108">値</span><span class="sxs-lookup"><span data-stu-id="8e966-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-110">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-110">1.0</span></span>|
|[<span data-ttu-id="8e966-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="8e966-112">Restricted</span></span>|
|[<span data-ttu-id="8e966-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-114">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8e966-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-115">Members and methods</span></span>

| <span data-ttu-id="8e966-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-116">Member</span></span> | <span data-ttu-id="8e966-117">種類</span><span class="sxs-lookup"><span data-stu-id="8e966-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8e966-118">attachments</span><span class="sxs-lookup"><span data-stu-id="8e966-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="8e966-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-119">Member</span></span> |
| [<span data-ttu-id="8e966-120">bcc</span><span class="sxs-lookup"><span data-stu-id="8e966-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8e966-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-121">Member</span></span> |
| [<span data-ttu-id="8e966-122">body</span><span class="sxs-lookup"><span data-stu-id="8e966-122">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="8e966-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-123">Member</span></span> |
| [<span data-ttu-id="8e966-124">cc</span><span class="sxs-lookup"><span data-stu-id="8e966-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8e966-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-125">Member</span></span> |
| [<span data-ttu-id="8e966-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="8e966-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="8e966-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-127">Member</span></span> |
| [<span data-ttu-id="8e966-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="8e966-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="8e966-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-129">Member</span></span> |
| [<span data-ttu-id="8e966-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="8e966-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="8e966-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-131">Member</span></span> |
| [<span data-ttu-id="8e966-132">end</span><span class="sxs-lookup"><span data-stu-id="8e966-132">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="8e966-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-133">Member</span></span> |
| [<span data-ttu-id="8e966-134">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="8e966-134">enhancedLocation</span></span>](#enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation) | <span data-ttu-id="8e966-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-135">Member</span></span> |
| [<span data-ttu-id="8e966-136">from</span><span class="sxs-lookup"><span data-stu-id="8e966-136">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="8e966-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-137">Member</span></span> |
| [<span data-ttu-id="8e966-138">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="8e966-138">internetHeaders</span></span>](#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders) | <span data-ttu-id="8e966-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-139">Member</span></span> |
| [<span data-ttu-id="8e966-140">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="8e966-140">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="8e966-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-141">Member</span></span> |
| [<span data-ttu-id="8e966-142">itemClass</span><span class="sxs-lookup"><span data-stu-id="8e966-142">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="8e966-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-143">Member</span></span> |
| [<span data-ttu-id="8e966-144">itemId</span><span class="sxs-lookup"><span data-stu-id="8e966-144">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="8e966-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-145">Member</span></span> |
| [<span data-ttu-id="8e966-146">itemType</span><span class="sxs-lookup"><span data-stu-id="8e966-146">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="8e966-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-147">Member</span></span> |
| [<span data-ttu-id="8e966-148">location</span><span class="sxs-lookup"><span data-stu-id="8e966-148">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="8e966-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-149">Member</span></span> |
| [<span data-ttu-id="8e966-150">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="8e966-150">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="8e966-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-151">Member</span></span> |
| [<span data-ttu-id="8e966-152">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="8e966-152">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="8e966-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-153">Member</span></span> |
| [<span data-ttu-id="8e966-154">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="8e966-154">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8e966-155">Member</span><span class="sxs-lookup"><span data-stu-id="8e966-155">Member</span></span> |
| [<span data-ttu-id="8e966-156">organizer</span><span class="sxs-lookup"><span data-stu-id="8e966-156">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="8e966-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-157">Member</span></span> |
| [<span data-ttu-id="8e966-158">recurrence</span><span class="sxs-lookup"><span data-stu-id="8e966-158">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="8e966-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-159">Member</span></span> |
| [<span data-ttu-id="8e966-160">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="8e966-160">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8e966-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-161">Member</span></span> |
| [<span data-ttu-id="8e966-162">sender</span><span class="sxs-lookup"><span data-stu-id="8e966-162">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="8e966-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-163">Member</span></span> |
| [<span data-ttu-id="8e966-164">seriesId</span><span class="sxs-lookup"><span data-stu-id="8e966-164">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="8e966-165">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-165">Member</span></span> |
| [<span data-ttu-id="8e966-166">start</span><span class="sxs-lookup"><span data-stu-id="8e966-166">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="8e966-167">Member</span><span class="sxs-lookup"><span data-stu-id="8e966-167">Member</span></span> |
| [<span data-ttu-id="8e966-168">subject</span><span class="sxs-lookup"><span data-stu-id="8e966-168">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="8e966-169">Member</span><span class="sxs-lookup"><span data-stu-id="8e966-169">Member</span></span> |
| [<span data-ttu-id="8e966-170">to</span><span class="sxs-lookup"><span data-stu-id="8e966-170">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="8e966-171">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-171">Member</span></span> |
| [<span data-ttu-id="8e966-172">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8e966-172">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="8e966-173">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-173">Method</span></span> |
| [<span data-ttu-id="8e966-174">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="8e966-174">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="8e966-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-175">Method</span></span> |
| [<span data-ttu-id="8e966-176">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="8e966-176">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="8e966-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-177">Method</span></span> |
| [<span data-ttu-id="8e966-178">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8e966-178">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="8e966-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-179">Method</span></span> |
| [<span data-ttu-id="8e966-180">close</span><span class="sxs-lookup"><span data-stu-id="8e966-180">close</span></span>](#close) | <span data-ttu-id="8e966-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-181">Method</span></span> |
| [<span data-ttu-id="8e966-182">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="8e966-182">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="8e966-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-183">Method</span></span> |
| [<span data-ttu-id="8e966-184">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="8e966-184">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="8e966-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-185">Method</span></span> |
| [<span data-ttu-id="8e966-186">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="8e966-186">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) | <span data-ttu-id="8e966-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-187">Method</span></span> |
| [<span data-ttu-id="8e966-188">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="8e966-188">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="8e966-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-189">Method</span></span> |
| [<span data-ttu-id="8e966-190">getEntities</span><span class="sxs-lookup"><span data-stu-id="8e966-190">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="8e966-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-191">Method</span></span> |
| [<span data-ttu-id="8e966-192">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="8e966-192">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="8e966-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-193">Method</span></span> |
| [<span data-ttu-id="8e966-194">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="8e966-194">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="8e966-195">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-195">Method</span></span> |
| [<span data-ttu-id="8e966-196">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="8e966-196">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="8e966-197">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-197">Method</span></span> |
| [<span data-ttu-id="8e966-198">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="8e966-198">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="8e966-199">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-199">Method</span></span> |
| [<span data-ttu-id="8e966-200">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="8e966-200">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="8e966-201">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-201">Method</span></span> |
| [<span data-ttu-id="8e966-202">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="8e966-202">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="8e966-203">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-203">Method</span></span> |
| [<span data-ttu-id="8e966-204">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="8e966-204">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="8e966-205">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-205">Method</span></span> |
| [<span data-ttu-id="8e966-206">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="8e966-206">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="8e966-207">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-207">Method</span></span> |
| [<span data-ttu-id="8e966-208">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="8e966-208">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="8e966-209">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-209">Method</span></span> |
| [<span data-ttu-id="8e966-210">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="8e966-210">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="8e966-211">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-211">Method</span></span> |
| [<span data-ttu-id="8e966-212">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8e966-212">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="8e966-213">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-213">Method</span></span> |
| [<span data-ttu-id="8e966-214">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="8e966-214">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="8e966-215">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-215">Method</span></span> |
| [<span data-ttu-id="8e966-216">saveAsync</span><span class="sxs-lookup"><span data-stu-id="8e966-216">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="8e966-217">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-217">Method</span></span> |
| [<span data-ttu-id="8e966-218">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="8e966-218">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="8e966-219">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-219">Method</span></span> |

### <a name="example"></a><span data-ttu-id="8e966-220">例</span><span class="sxs-lookup"><span data-stu-id="8e966-220">Example</span></span>

<span data-ttu-id="8e966-221">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8e966-221">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="8e966-222">メンバー</span><span class="sxs-lookup"><span data-stu-id="8e966-222">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="8e966-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="8e966-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="8e966-224">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="8e966-224">Gets the item's attachments as an array.</span></span> <span data-ttu-id="8e966-225">閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8e966-225">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8e966-226">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="8e966-226">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="8e966-227">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8e966-227">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-228">種類:</span><span class="sxs-lookup"><span data-stu-id="8e966-228">Type:</span></span>

*   <span data-ttu-id="8e966-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="8e966-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-230">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-230">Requirements</span></span>

|<span data-ttu-id="8e966-231">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-231">Requirement</span></span>|<span data-ttu-id="8e966-232">値</span><span class="sxs-lookup"><span data-stu-id="8e966-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-233">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-234">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-234">1.0</span></span>|
|[<span data-ttu-id="8e966-235">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-235">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-236">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-236">ReadItem</span></span>|
|[<span data-ttu-id="8e966-237">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-237">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-238">読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-238">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-239">例</span><span class="sxs-lookup"><span data-stu-id="8e966-239">Example</span></span>

<span data-ttu-id="8e966-240">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="8e966-240">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8e966-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8e966-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8e966-242">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="8e966-242">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="8e966-243">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8e966-243">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-244">種類:</span><span class="sxs-lookup"><span data-stu-id="8e966-244">Type:</span></span>

*   [<span data-ttu-id="8e966-245">Recipients</span><span class="sxs-lookup"><span data-stu-id="8e966-245">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="8e966-246">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-246">Requirements</span></span>

|<span data-ttu-id="8e966-247">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-247">Requirement</span></span>|<span data-ttu-id="8e966-248">値</span><span class="sxs-lookup"><span data-stu-id="8e966-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-249">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-250">1.1</span><span class="sxs-lookup"><span data-stu-id="8e966-250">1.1</span></span>|
|[<span data-ttu-id="8e966-251">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-251">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-252">ReadItem</span></span>|
|[<span data-ttu-id="8e966-253">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-253">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-254">作成</span><span class="sxs-lookup"><span data-stu-id="8e966-254">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-255">例</span><span class="sxs-lookup"><span data-stu-id="8e966-255">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="8e966-256">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="8e966-256">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="8e966-257">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="8e966-257">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-258">型:</span><span class="sxs-lookup"><span data-stu-id="8e966-258">Type:</span></span>

*   [<span data-ttu-id="8e966-259">Body</span><span class="sxs-lookup"><span data-stu-id="8e966-259">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="8e966-260">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-260">Requirements</span></span>

|<span data-ttu-id="8e966-261">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-261">Requirement</span></span>|<span data-ttu-id="8e966-262">値</span><span class="sxs-lookup"><span data-stu-id="8e966-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-263">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-264">1.1</span><span class="sxs-lookup"><span data-stu-id="8e966-264">1.1</span></span>|
|[<span data-ttu-id="8e966-265">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-265">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-266">ReadItem</span></span>|
|[<span data-ttu-id="8e966-267">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-267">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-268">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-268">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8e966-269">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8e966-269">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8e966-270">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8e966-270">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="8e966-271">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="8e966-271">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8e966-272">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8e966-272">Read mode</span></span>

<span data-ttu-id="8e966-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8e966-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8e966-275">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8e966-275">Compose mode</span></span>

<span data-ttu-id="8e966-276">`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-276">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-277">種類:</span><span class="sxs-lookup"><span data-stu-id="8e966-277">Type:</span></span>

*   <span data-ttu-id="8e966-278">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8e966-278">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-279">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-279">Requirements</span></span>

|<span data-ttu-id="8e966-280">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-280">Requirement</span></span>|<span data-ttu-id="8e966-281">値</span><span class="sxs-lookup"><span data-stu-id="8e966-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-282">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-283">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-283">1.0</span></span>|
|[<span data-ttu-id="8e966-284">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-284">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-285">ReadItem</span></span>|
|[<span data-ttu-id="8e966-286">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-286">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-287">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-287">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-288">例</span><span class="sxs-lookup"><span data-stu-id="8e966-288">Example</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="8e966-289">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="8e966-289">(nullable) conversationId :String</span></span>

<span data-ttu-id="8e966-290">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="8e966-290">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="8e966-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="8e966-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="8e966-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-295">型:</span><span class="sxs-lookup"><span data-stu-id="8e966-295">Type:</span></span>

*   <span data-ttu-id="8e966-296">String</span><span class="sxs-lookup"><span data-stu-id="8e966-296">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-297">Requirements</span><span class="sxs-lookup"><span data-stu-id="8e966-297">Requirements</span></span>

|<span data-ttu-id="8e966-298">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-298">Requirement</span></span>|<span data-ttu-id="8e966-299">値</span><span class="sxs-lookup"><span data-stu-id="8e966-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-300">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-300">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-301">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-301">1.0</span></span>|
|[<span data-ttu-id="8e966-302">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-302">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-303">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-303">ReadItem</span></span>|
|[<span data-ttu-id="8e966-304">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-304">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-305">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8e966-305">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="8e966-306">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="8e966-306">dateTimeCreated :Date</span></span>

<span data-ttu-id="8e966-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8e966-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-309">種類:</span><span class="sxs-lookup"><span data-stu-id="8e966-309">Type:</span></span>

*   <span data-ttu-id="8e966-310">日付</span><span class="sxs-lookup"><span data-stu-id="8e966-310">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-311">Requirements</span><span class="sxs-lookup"><span data-stu-id="8e966-311">Requirements</span></span>

|<span data-ttu-id="8e966-312">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-312">Requirement</span></span>|<span data-ttu-id="8e966-313">値</span><span class="sxs-lookup"><span data-stu-id="8e966-313">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-314">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-314">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-315">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-315">1.0</span></span>|
|[<span data-ttu-id="8e966-316">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-316">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-317">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-317">ReadItem</span></span>|
|[<span data-ttu-id="8e966-318">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-318">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-319">読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-319">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-320">例</span><span class="sxs-lookup"><span data-stu-id="8e966-320">Example</span></span>

```javascript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="8e966-321">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="8e966-321">dateTimeModified :Date</span></span>

<span data-ttu-id="8e966-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8e966-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8e966-324">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8e966-324">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-325">型:</span><span class="sxs-lookup"><span data-stu-id="8e966-325">Type:</span></span>

*   <span data-ttu-id="8e966-326">日付</span><span class="sxs-lookup"><span data-stu-id="8e966-326">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-327">Requirements</span><span class="sxs-lookup"><span data-stu-id="8e966-327">Requirements</span></span>

|<span data-ttu-id="8e966-328">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-328">Requirement</span></span>|<span data-ttu-id="8e966-329">値</span><span class="sxs-lookup"><span data-stu-id="8e966-329">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-330">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-330">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-331">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-331">1.0</span></span>|
|[<span data-ttu-id="8e966-332">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-332">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-333">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-333">ReadItem</span></span>|
|[<span data-ttu-id="8e966-334">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-334">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-335">読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-335">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-336">例</span><span class="sxs-lookup"><span data-stu-id="8e966-336">Example</span></span>

```javascript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="8e966-337">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="8e966-337">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="8e966-338">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8e966-338">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="8e966-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="8e966-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8e966-341">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8e966-341">Read mode</span></span>

<span data-ttu-id="8e966-342">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-342">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8e966-343">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8e966-343">Compose mode</span></span>

<span data-ttu-id="8e966-344">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-344">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="8e966-345">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8e966-345">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-346">型:</span><span class="sxs-lookup"><span data-stu-id="8e966-346">Type:</span></span>

*   <span data-ttu-id="8e966-347">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="8e966-347">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-348">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-348">Requirements</span></span>

|<span data-ttu-id="8e966-349">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-349">Requirement</span></span>|<span data-ttu-id="8e966-350">値</span><span class="sxs-lookup"><span data-stu-id="8e966-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-351">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-352">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-352">1.0</span></span>|
|[<span data-ttu-id="8e966-353">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-353">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-354">ReadItem</span></span>|
|[<span data-ttu-id="8e966-355">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-355">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-356">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-356">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-357">例</span><span class="sxs-lookup"><span data-stu-id="8e966-357">Example</span></span>

<span data-ttu-id="8e966-358">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="8e966-358">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="8e966-359">enhancedLocation:[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="8e966-359">enhancedLocation :[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="8e966-360">取得または予定の場所を設定します。</span><span class="sxs-lookup"><span data-stu-id="8e966-360">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8e966-361">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8e966-361">Read mode</span></span>

<span data-ttu-id="8e966-362">`enhancedLocation`を使用すると、予定に関連付けられている (それぞれは、 [LocationDetails](/javascript/api/outlook/office.locationdetails)オブジェクトで表されます) の場所のセットを取得する[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-362">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8e966-363">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8e966-363">Compose mode</span></span>

<span data-ttu-id="8e966-364">`enhancedLocation`を取得、削除、または予定の場所を追加するメソッドを提供する[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-364">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-365">種類:</span><span class="sxs-lookup"><span data-stu-id="8e966-365">Type:</span></span>

*   [<span data-ttu-id="8e966-366">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="8e966-366">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="8e966-367">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-367">Requirements</span></span>

|<span data-ttu-id="8e966-368">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-368">Requirement</span></span>|<span data-ttu-id="8e966-369">値</span><span class="sxs-lookup"><span data-stu-id="8e966-369">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-370">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-370">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-371">プレビュー</span><span class="sxs-lookup"><span data-stu-id="8e966-371">Preview</span></span>|
|[<span data-ttu-id="8e966-372">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-372">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-373">ReadItem</span></span>|
|[<span data-ttu-id="8e966-374">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-374">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-375">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-375">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-376">例</span><span class="sxs-lookup"><span data-stu-id="8e966-376">Example</span></span>

<span data-ttu-id="8e966-377">次の例では、予定に関連付けられている現在の場所を取得します。</span><span class="sxs-lookup"><span data-stu-id="8e966-377">The following example gets the current locations associated with the appointment.</span></span>

```javascript
Office.context.mailbox.item.enhancedLocation.getAsync(callbackFunction);

function callbackFunction(asyncResult) {
  asyncResult.value.forEach(function (place) {
    console.log("Display name: " + place.displayName);
    console.log("Type: " + place.locationIdentifier.type);
    if (place.locationIdentifier.type == Office.MailboxEnums.LocationType.Room) {
      console.log("Email address: " + place.emailAddress);
    }
  });
}

// Sample output:
// Display name: Conf Room 14
// Type: room
// Email address: cr14@contoso.com
// Display name: Paris
// Type: custom
```

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="8e966-378">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="8e966-378">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="8e966-379">メッセージの送信者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="8e966-379">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="8e966-p112">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="8e966-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8e966-382">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="8e966-382">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8e966-383">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8e966-383">Read mode</span></span>

<span data-ttu-id="8e966-384">`from` プロパティは `EmailAddressDetails` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-384">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="8e966-385">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8e966-385">Compose mode</span></span>

<span data-ttu-id="8e966-386">`from` プロパティは From 値を取得するメソッドを提供する `From` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-386">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8e966-387">種類:</span><span class="sxs-lookup"><span data-stu-id="8e966-387">Type:</span></span>

*   <span data-ttu-id="8e966-388">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="8e966-388">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-389">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-389">Requirements</span></span>

|<span data-ttu-id="8e966-390">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-390">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="8e966-391">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-391">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-392">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-392">1.0</span></span>|<span data-ttu-id="8e966-393">1.7</span><span class="sxs-lookup"><span data-stu-id="8e966-393">1.7</span></span>|
|[<span data-ttu-id="8e966-394">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-394">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-395">ReadItem</span></span>|<span data-ttu-id="8e966-396">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8e966-396">ReadWriteItem</span></span>|
|[<span data-ttu-id="8e966-397">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-397">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-398">Read</span><span class="sxs-lookup"><span data-stu-id="8e966-398">Read</span></span>|<span data-ttu-id="8e966-399">Compose</span><span class="sxs-lookup"><span data-stu-id="8e966-399">Compose</span></span>|

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="8e966-400">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="8e966-400">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="8e966-401">メッセージのインターネット ヘッダーを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8e966-401">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-402">型:</span><span class="sxs-lookup"><span data-stu-id="8e966-402">Type:</span></span>

*   [<span data-ttu-id="8e966-403">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="8e966-403">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="8e966-404">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-404">Requirements</span></span>

|<span data-ttu-id="8e966-405">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-405">Requirement</span></span>|<span data-ttu-id="8e966-406">値</span><span class="sxs-lookup"><span data-stu-id="8e966-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-407">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-408">プレビュー</span><span class="sxs-lookup"><span data-stu-id="8e966-408">Preview</span></span>|
|[<span data-ttu-id="8e966-409">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-409">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-410">ReadItem</span></span>|
|[<span data-ttu-id="8e966-411">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-411">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-412">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="8e966-412">Compose or read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="8e966-413">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="8e966-413">internetMessageId :String</span></span>

<span data-ttu-id="8e966-p113">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8e966-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-416">種類:</span><span class="sxs-lookup"><span data-stu-id="8e966-416">Type:</span></span>

*   <span data-ttu-id="8e966-417">String</span><span class="sxs-lookup"><span data-stu-id="8e966-417">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-418">Requirements</span><span class="sxs-lookup"><span data-stu-id="8e966-418">Requirements</span></span>

|<span data-ttu-id="8e966-419">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-419">Requirement</span></span>|<span data-ttu-id="8e966-420">値</span><span class="sxs-lookup"><span data-stu-id="8e966-420">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-421">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-421">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-422">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-422">1.0</span></span>|
|[<span data-ttu-id="8e966-423">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-423">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-424">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-424">ReadItem</span></span>|
|[<span data-ttu-id="8e966-425">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-425">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-426">読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-426">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-427">例</span><span class="sxs-lookup"><span data-stu-id="8e966-427">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="8e966-428">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="8e966-428">itemClass :String</span></span>

<span data-ttu-id="8e966-p114">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8e966-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="8e966-p115">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="8e966-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="8e966-433">型</span><span class="sxs-lookup"><span data-stu-id="8e966-433">Type</span></span>|<span data-ttu-id="8e966-434">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-434">Description</span></span>|<span data-ttu-id="8e966-435">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="8e966-435">item class</span></span>|
|---|---|---|
|<span data-ttu-id="8e966-436">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="8e966-436">Appointment items</span></span>|<span data-ttu-id="8e966-437">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8e966-437">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="8e966-438">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="8e966-438">Message items</span></span>|<span data-ttu-id="8e966-439">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="8e966-439">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="8e966-440">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="8e966-440">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-441">種類:</span><span class="sxs-lookup"><span data-stu-id="8e966-441">Type:</span></span>

*   <span data-ttu-id="8e966-442">String</span><span class="sxs-lookup"><span data-stu-id="8e966-442">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-443">Requirements</span><span class="sxs-lookup"><span data-stu-id="8e966-443">Requirements</span></span>

|<span data-ttu-id="8e966-444">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-444">Requirement</span></span>|<span data-ttu-id="8e966-445">値</span><span class="sxs-lookup"><span data-stu-id="8e966-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-446">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-447">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-447">1.0</span></span>|
|[<span data-ttu-id="8e966-448">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-449">ReadItem</span></span>|
|[<span data-ttu-id="8e966-450">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-451">読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-452">例</span><span class="sxs-lookup"><span data-stu-id="8e966-452">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="8e966-453">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="8e966-453">(nullable) itemId :String</span></span>

<span data-ttu-id="8e966-p116">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8e966-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8e966-456">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="8e966-456">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="8e966-457">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="8e966-457">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="8e966-458">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8e966-458">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="8e966-459">詳細は、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8e966-459">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="8e966-p118">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-462">種類:</span><span class="sxs-lookup"><span data-stu-id="8e966-462">Type:</span></span>

*   <span data-ttu-id="8e966-463">String</span><span class="sxs-lookup"><span data-stu-id="8e966-463">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-464">Requirements</span><span class="sxs-lookup"><span data-stu-id="8e966-464">Requirements</span></span>

|<span data-ttu-id="8e966-465">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-465">Requirement</span></span>|<span data-ttu-id="8e966-466">値</span><span class="sxs-lookup"><span data-stu-id="8e966-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-467">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-468">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-468">1.0</span></span>|
|[<span data-ttu-id="8e966-469">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-470">ReadItem</span></span>|
|[<span data-ttu-id="8e966-471">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-472">読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-473">例</span><span class="sxs-lookup"><span data-stu-id="8e966-473">Example</span></span>

<span data-ttu-id="8e966-p119">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="8e966-476">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="8e966-476">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="8e966-477">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="8e966-477">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="8e966-478">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="8e966-478">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-479">型:</span><span class="sxs-lookup"><span data-stu-id="8e966-479">Type:</span></span>

*   [<span data-ttu-id="8e966-480">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="8e966-480">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="8e966-481">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-481">Requirements</span></span>

|<span data-ttu-id="8e966-482">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-482">Requirement</span></span>|<span data-ttu-id="8e966-483">値</span><span class="sxs-lookup"><span data-stu-id="8e966-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-484">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-485">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-485">1.0</span></span>|
|[<span data-ttu-id="8e966-486">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-486">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-487">ReadItem</span></span>|
|[<span data-ttu-id="8e966-488">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-488">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-489">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-489">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-490">例</span><span class="sxs-lookup"><span data-stu-id="8e966-490">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="8e966-491">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="8e966-491">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="8e966-492">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8e966-492">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8e966-493">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8e966-493">Read mode</span></span>

<span data-ttu-id="8e966-494">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-494">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8e966-495">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8e966-495">Compose mode</span></span>

<span data-ttu-id="8e966-496">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-496">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-497">種類:</span><span class="sxs-lookup"><span data-stu-id="8e966-497">Type:</span></span>

*   <span data-ttu-id="8e966-498">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="8e966-498">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-499">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-499">Requirements</span></span>

|<span data-ttu-id="8e966-500">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-500">Requirement</span></span>|<span data-ttu-id="8e966-501">値</span><span class="sxs-lookup"><span data-stu-id="8e966-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-502">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-503">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-503">1.0</span></span>|
|[<span data-ttu-id="8e966-504">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-505">ReadItem</span></span>|
|[<span data-ttu-id="8e966-506">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-507">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-507">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-508">例</span><span class="sxs-lookup"><span data-stu-id="8e966-508">Example</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="8e966-509">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="8e966-509">normalizedSubject :String</span></span>

<span data-ttu-id="8e966-p120">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8e966-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="8e966-p121">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="8e966-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-514">型:</span><span class="sxs-lookup"><span data-stu-id="8e966-514">Type:</span></span>

*   <span data-ttu-id="8e966-515">String</span><span class="sxs-lookup"><span data-stu-id="8e966-515">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-516">Requirements</span><span class="sxs-lookup"><span data-stu-id="8e966-516">Requirements</span></span>

|<span data-ttu-id="8e966-517">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-517">Requirement</span></span>|<span data-ttu-id="8e966-518">値</span><span class="sxs-lookup"><span data-stu-id="8e966-518">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-519">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-519">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-520">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-520">1.0</span></span>|
|[<span data-ttu-id="8e966-521">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-521">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-522">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-522">ReadItem</span></span>|
|[<span data-ttu-id="8e966-523">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-523">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-524">読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-524">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-525">例</span><span class="sxs-lookup"><span data-stu-id="8e966-525">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="8e966-526">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="8e966-526">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="8e966-527">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="8e966-527">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-528">型:</span><span class="sxs-lookup"><span data-stu-id="8e966-528">Type:</span></span>

*   [<span data-ttu-id="8e966-529">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="8e966-529">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="8e966-530">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-530">Requirements</span></span>

|<span data-ttu-id="8e966-531">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-531">Requirement</span></span>|<span data-ttu-id="8e966-532">値</span><span class="sxs-lookup"><span data-stu-id="8e966-532">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-533">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-533">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-534">1.3</span><span class="sxs-lookup"><span data-stu-id="8e966-534">1.3</span></span>|
|[<span data-ttu-id="8e966-535">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-535">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-536">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-536">ReadItem</span></span>|
|[<span data-ttu-id="8e966-537">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-537">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-538">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-538">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8e966-539">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8e966-539">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8e966-540">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8e966-540">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="8e966-541">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="8e966-541">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8e966-542">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8e966-542">Read mode</span></span>

<span data-ttu-id="8e966-543">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-543">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8e966-544">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8e966-544">Compose mode</span></span>

<span data-ttu-id="8e966-545">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-545">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-546">種類:</span><span class="sxs-lookup"><span data-stu-id="8e966-546">Type:</span></span>

*   <span data-ttu-id="8e966-547">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8e966-547">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-548">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-548">Requirements</span></span>

|<span data-ttu-id="8e966-549">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-549">Requirement</span></span>|<span data-ttu-id="8e966-550">値</span><span class="sxs-lookup"><span data-stu-id="8e966-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-551">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-552">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-552">1.0</span></span>|
|[<span data-ttu-id="8e966-553">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-554">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-554">ReadItem</span></span>|
|[<span data-ttu-id="8e966-555">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-556">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-556">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-557">例</span><span class="sxs-lookup"><span data-stu-id="8e966-557">Example</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="8e966-558">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="8e966-558">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="8e966-559">指定の会議の開催者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="8e966-559">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8e966-560">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8e966-560">Read mode</span></span>

<span data-ttu-id="8e966-561">`organizer` プロパティは、会議開催者を表す [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-561">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8e966-562">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8e966-562">Compose mode</span></span>

<span data-ttu-id="8e966-563">`organizer` プロパティは Organizer 値を取得するメソッドを提供する [Organizer](/javascript/api/outlook/office.organizer) オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-563">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-564">種類:</span><span class="sxs-lookup"><span data-stu-id="8e966-564">Type:</span></span>

*   <span data-ttu-id="8e966-565">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="8e966-565">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-566">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-566">Requirements</span></span>

|<span data-ttu-id="8e966-567">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-567">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="8e966-568">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-568">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-569">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-569">1.0</span></span>|<span data-ttu-id="8e966-570">1.7</span><span class="sxs-lookup"><span data-stu-id="8e966-570">1.7</span></span>|
|[<span data-ttu-id="8e966-571">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-571">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-572">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-572">ReadItem</span></span>|<span data-ttu-id="8e966-573">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8e966-573">ReadWriteItem</span></span>|
|[<span data-ttu-id="8e966-574">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-574">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-575">Read</span><span class="sxs-lookup"><span data-stu-id="8e966-575">Read</span></span>|<span data-ttu-id="8e966-576">Compose</span><span class="sxs-lookup"><span data-stu-id="8e966-576">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-577">例</span><span class="sxs-lookup"><span data-stu-id="8e966-577">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="8e966-578">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="8e966-578">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="8e966-579">予定の定期的なパターンを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8e966-579">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="8e966-580">会議出席依頼の定期的なパターンを取得します。</span><span class="sxs-lookup"><span data-stu-id="8e966-580">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="8e966-581">予定アイテムの閲覧モードと新規作成モード。</span><span class="sxs-lookup"><span data-stu-id="8e966-581">Read and compose modes for appointment items.</span></span> <span data-ttu-id="8e966-582">会議出席依頼アイテムの閲覧モード。</span><span class="sxs-lookup"><span data-stu-id="8e966-582">Read mode for meeting request items.</span></span>

<span data-ttu-id="8e966-583">`recurrence` プロパティは、アイテムがシリーズか、シリーズに含まれるインスタンスの場合、定期的な予定または会議出席依頼に対して [recurrence](/javascript/api/outlook/office.recurrence) オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-583">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="8e966-584">`null` は、単発の予定および単発の予定の会議出席依頼に対して返されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-584">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="8e966-585">`undefined` は、会議出席依頼ではないメッセージに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-585">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="8e966-586">注: 会議出席依頼の `itemClass` 値は IPM.Schedule.Meeting.Request です。</span><span class="sxs-lookup"><span data-stu-id="8e966-586">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="8e966-587">注: recurrence オブジェクトが `null` の場合、オブジェクトがシリーズの一部ではなく、1 つの単発の予定または 1 つの単発の予定の会議出席依頼であることを示します。</span><span class="sxs-lookup"><span data-stu-id="8e966-587">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-588">型:</span><span class="sxs-lookup"><span data-stu-id="8e966-588">Type:</span></span>

* [<span data-ttu-id="8e966-589">Recurrence</span><span class="sxs-lookup"><span data-stu-id="8e966-589">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="8e966-590">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-590">Requirement</span></span>|<span data-ttu-id="8e966-591">値</span><span class="sxs-lookup"><span data-stu-id="8e966-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-592">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-593">1.7</span><span class="sxs-lookup"><span data-stu-id="8e966-593">1.7</span></span>|
|[<span data-ttu-id="8e966-594">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-594">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-595">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-595">ReadItem</span></span>|
|[<span data-ttu-id="8e966-596">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-596">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-597">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-597">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8e966-598">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8e966-598">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8e966-599">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8e966-599">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="8e966-600">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="8e966-600">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8e966-601">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8e966-601">Read mode</span></span>

<span data-ttu-id="8e966-602">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-602">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8e966-603">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8e966-603">Compose mode</span></span>

<span data-ttu-id="8e966-604">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-604">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-605">種類:</span><span class="sxs-lookup"><span data-stu-id="8e966-605">Type:</span></span>

*   <span data-ttu-id="8e966-606">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8e966-606">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-607">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-607">Requirements</span></span>

|<span data-ttu-id="8e966-608">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-608">Requirement</span></span>|<span data-ttu-id="8e966-609">値</span><span class="sxs-lookup"><span data-stu-id="8e966-609">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-610">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-610">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-611">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-611">1.0</span></span>|
|[<span data-ttu-id="8e966-612">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-612">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-613">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-613">ReadItem</span></span>|
|[<span data-ttu-id="8e966-614">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-614">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-615">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-615">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-616">例</span><span class="sxs-lookup"><span data-stu-id="8e966-616">Example</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="8e966-617">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="8e966-617">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="8e966-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8e966-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="8e966-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="8e966-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8e966-622">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="8e966-622">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-623">種類:</span><span class="sxs-lookup"><span data-stu-id="8e966-623">Type:</span></span>

*   [<span data-ttu-id="8e966-624">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8e966-624">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="8e966-625">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-625">Requirements</span></span>

|<span data-ttu-id="8e966-626">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-626">Requirement</span></span>|<span data-ttu-id="8e966-627">値</span><span class="sxs-lookup"><span data-stu-id="8e966-627">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-628">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-628">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-629">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-629">1.0</span></span>|
|[<span data-ttu-id="8e966-630">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-630">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-631">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-631">ReadItem</span></span>|
|[<span data-ttu-id="8e966-632">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-632">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-633">読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-633">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-634">例</span><span class="sxs-lookup"><span data-stu-id="8e966-634">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="8e966-635">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="8e966-635">(nullable) seriesId :String</span></span>

<span data-ttu-id="8e966-636">あるインスタンスが属するシリーズの ID を取得します。</span><span class="sxs-lookup"><span data-stu-id="8e966-636">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="8e966-637">OWA と Outlook では、`seriesId` はこのアイテムが属する親 (シリーズ) アイテムの Exchange Web Services (EWS) ID を返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-637">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="8e966-638">ただし、iOS と Android の場合、`seriesId` は親アイテムの REST ID を返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-638">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="8e966-639">`seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="8e966-639">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="8e966-640">`seriesId` プロパティは、Outlook REST API で使用される Outlook ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="8e966-640">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="8e966-641">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8e966-641">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="8e966-642">詳細については、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8e966-642">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="8e966-643">`seriesId` プロパティは、単発の予定、シリーズ アイテム、会議出席依頼など、親アイテムを持たないアイテムに対して `null` を返し、会議出席依頼ではないその他のアイテムに対して `undefined` を返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-643">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-644">型:</span><span class="sxs-lookup"><span data-stu-id="8e966-644">Type:</span></span>

* <span data-ttu-id="8e966-645">String</span><span class="sxs-lookup"><span data-stu-id="8e966-645">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-646">Requirements</span><span class="sxs-lookup"><span data-stu-id="8e966-646">Requirements</span></span>

|<span data-ttu-id="8e966-647">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-647">Requirement</span></span>|<span data-ttu-id="8e966-648">値</span><span class="sxs-lookup"><span data-stu-id="8e966-648">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-649">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-649">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-650">1.7</span><span class="sxs-lookup"><span data-stu-id="8e966-650">1.7</span></span>|
|[<span data-ttu-id="8e966-651">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-651">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-652">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-652">ReadItem</span></span>|
|[<span data-ttu-id="8e966-653">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-653">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-654">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-654">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-655">例</span><span class="sxs-lookup"><span data-stu-id="8e966-655">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="8e966-656">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="8e966-656">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="8e966-657">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8e966-657">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="8e966-p130">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="8e966-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8e966-660">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8e966-660">Read mode</span></span>

<span data-ttu-id="8e966-661">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-661">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8e966-662">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8e966-662">Compose mode</span></span>

<span data-ttu-id="8e966-663">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-663">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="8e966-664">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8e966-664">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-665">種類:</span><span class="sxs-lookup"><span data-stu-id="8e966-665">Type:</span></span>

*   <span data-ttu-id="8e966-666">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="8e966-666">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-667">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-667">Requirements</span></span>

|<span data-ttu-id="8e966-668">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-668">Requirement</span></span>|<span data-ttu-id="8e966-669">値</span><span class="sxs-lookup"><span data-stu-id="8e966-669">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-670">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-670">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-671">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-671">1.0</span></span>|
|[<span data-ttu-id="8e966-672">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-672">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-673">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-673">ReadItem</span></span>|
|[<span data-ttu-id="8e966-674">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-674">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-675">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-675">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-676">例</span><span class="sxs-lookup"><span data-stu-id="8e966-676">Example</span></span>

<span data-ttu-id="8e966-677">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="8e966-677">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="8e966-678">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="8e966-678">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="8e966-679">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8e966-679">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="8e966-680">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8e966-680">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8e966-681">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8e966-681">Read mode</span></span>

<span data-ttu-id="8e966-p131">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="8e966-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="8e966-684">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8e966-684">Compose mode</span></span>

<span data-ttu-id="8e966-685">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-685">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8e966-686">型:</span><span class="sxs-lookup"><span data-stu-id="8e966-686">Type:</span></span>

*   <span data-ttu-id="8e966-687">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="8e966-687">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-688">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-688">Requirements</span></span>

|<span data-ttu-id="8e966-689">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-689">Requirement</span></span>|<span data-ttu-id="8e966-690">値</span><span class="sxs-lookup"><span data-stu-id="8e966-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-691">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-692">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-692">1.0</span></span>|
|[<span data-ttu-id="8e966-693">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-693">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-694">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-694">ReadItem</span></span>|
|[<span data-ttu-id="8e966-695">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-695">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-696">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-696">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="8e966-697">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8e966-697">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="8e966-698">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8e966-698">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="8e966-699">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="8e966-699">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8e966-700">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8e966-700">Read mode</span></span>

<span data-ttu-id="8e966-p133">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8e966-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8e966-703">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8e966-703">Compose mode</span></span>

<span data-ttu-id="8e966-704">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-704">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="8e966-705">種類:</span><span class="sxs-lookup"><span data-stu-id="8e966-705">Type:</span></span>

*   <span data-ttu-id="8e966-706">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8e966-706">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-707">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-707">Requirements</span></span>

|<span data-ttu-id="8e966-708">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-708">Requirement</span></span>|<span data-ttu-id="8e966-709">値</span><span class="sxs-lookup"><span data-stu-id="8e966-709">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-710">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-710">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-711">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-711">1.0</span></span>|
|[<span data-ttu-id="8e966-712">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-712">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-713">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-713">ReadItem</span></span>|
|[<span data-ttu-id="8e966-714">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-714">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-715">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-715">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-716">例</span><span class="sxs-lookup"><span data-stu-id="8e966-716">Example</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="8e966-717">メソッド</span><span class="sxs-lookup"><span data-stu-id="8e966-717">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="8e966-718">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8e966-718">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8e966-719">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="8e966-719">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="8e966-720">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="8e966-720">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="8e966-721">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="8e966-721">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e966-722">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8e966-722">Parameters:</span></span>
|<span data-ttu-id="8e966-723">名前</span><span class="sxs-lookup"><span data-stu-id="8e966-723">Name</span></span>|<span data-ttu-id="8e966-724">型</span><span class="sxs-lookup"><span data-stu-id="8e966-724">Type</span></span>|<span data-ttu-id="8e966-725">属性</span><span class="sxs-lookup"><span data-stu-id="8e966-725">Attributes</span></span>|<span data-ttu-id="8e966-726">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-726">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="8e966-727">String</span><span class="sxs-lookup"><span data-stu-id="8e966-727">String</span></span>||<span data-ttu-id="8e966-p134">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="8e966-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="8e966-730">String</span><span class="sxs-lookup"><span data-stu-id="8e966-730">String</span></span>||<span data-ttu-id="8e966-p135">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="8e966-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="8e966-733">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-733">Object</span></span>|<span data-ttu-id="8e966-734">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-734">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-735">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8e966-735">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8e966-736">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-736">Object</span></span>|<span data-ttu-id="8e966-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-737">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-738">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8e966-738">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="8e966-739">Boolean</span><span class="sxs-lookup"><span data-stu-id="8e966-739">Boolean</span></span>|<span data-ttu-id="8e966-740">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-740">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-741">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="8e966-741">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="8e966-742">function</span><span class="sxs-lookup"><span data-stu-id="8e966-742">function</span></span>|<span data-ttu-id="8e966-743">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-743">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-744">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-744">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8e966-745">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-745">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8e966-746">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="8e966-746">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8e966-747">エラー</span><span class="sxs-lookup"><span data-stu-id="8e966-747">Errors</span></span>

|<span data-ttu-id="8e966-748">エラー コード</span><span class="sxs-lookup"><span data-stu-id="8e966-748">Error code</span></span>|<span data-ttu-id="8e966-749">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-749">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="8e966-750">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="8e966-750">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="8e966-751">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="8e966-751">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="8e966-752">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="8e966-752">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e966-753">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-753">Requirements</span></span>

|<span data-ttu-id="8e966-754">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-754">Requirement</span></span>|<span data-ttu-id="8e966-755">値</span><span class="sxs-lookup"><span data-stu-id="8e966-755">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-756">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-756">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-757">1.1</span><span class="sxs-lookup"><span data-stu-id="8e966-757">1.1</span></span>|
|[<span data-ttu-id="8e966-758">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-758">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-759">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8e966-759">ReadWriteItem</span></span>|
|[<span data-ttu-id="8e966-760">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-760">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-761">作成</span><span class="sxs-lookup"><span data-stu-id="8e966-761">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8e966-762">例</span><span class="sxs-lookup"><span data-stu-id="8e966-762">Examples</span></span>

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

<span data-ttu-id="8e966-763">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="8e966-763">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="8e966-764">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8e966-764">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8e966-765">ファイルを添付ファイルとして base64 エンコーディングからメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="8e966-765">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="8e966-766">`addFileAttachmentFromBase64Async` メソッドは、base64 エンコーディングからファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="8e966-766">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="8e966-767">このメソッドによって、AsyncResult.value オブジェクトの添付ファイル識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-767">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="8e966-768">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="8e966-768">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e966-769">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8e966-769">Parameters:</span></span>
|<span data-ttu-id="8e966-770">名前</span><span class="sxs-lookup"><span data-stu-id="8e966-770">Name</span></span>|<span data-ttu-id="8e966-771">型</span><span class="sxs-lookup"><span data-stu-id="8e966-771">Type</span></span>|<span data-ttu-id="8e966-772">属性</span><span class="sxs-lookup"><span data-stu-id="8e966-772">Attributes</span></span>|<span data-ttu-id="8e966-773">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-773">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="8e966-774">String</span><span class="sxs-lookup"><span data-stu-id="8e966-774">String</span></span>||<span data-ttu-id="8e966-775">電子メールまたはイベントに追加する画像またはファイルの base64 でエンコードされたコンテンツ。</span><span class="sxs-lookup"><span data-stu-id="8e966-775">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="8e966-776">String</span><span class="sxs-lookup"><span data-stu-id="8e966-776">String</span></span>||<span data-ttu-id="8e966-p137">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="8e966-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="8e966-779">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-779">Object</span></span>|<span data-ttu-id="8e966-780">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-780">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-781">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8e966-781">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8e966-782">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-782">Object</span></span>|<span data-ttu-id="8e966-783">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-783">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-784">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8e966-784">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="8e966-785">Boolean</span><span class="sxs-lookup"><span data-stu-id="8e966-785">Boolean</span></span>|<span data-ttu-id="8e966-786">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-786">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-787">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="8e966-787">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="8e966-788">function</span><span class="sxs-lookup"><span data-stu-id="8e966-788">function</span></span>|<span data-ttu-id="8e966-789">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-789">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-790">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-790">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8e966-791">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-791">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8e966-792">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="8e966-792">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8e966-793">エラー</span><span class="sxs-lookup"><span data-stu-id="8e966-793">Errors</span></span>

|<span data-ttu-id="8e966-794">エラー コード</span><span class="sxs-lookup"><span data-stu-id="8e966-794">Error code</span></span>|<span data-ttu-id="8e966-795">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-795">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="8e966-796">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="8e966-796">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="8e966-797">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="8e966-797">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="8e966-798">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="8e966-798">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e966-799">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-799">Requirements</span></span>

|<span data-ttu-id="8e966-800">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-800">Requirement</span></span>|<span data-ttu-id="8e966-801">値</span><span class="sxs-lookup"><span data-stu-id="8e966-801">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-802">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-802">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-803">プレビュー</span><span class="sxs-lookup"><span data-stu-id="8e966-803">Preview</span></span>|
|[<span data-ttu-id="8e966-804">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-804">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-805">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8e966-805">ReadWriteItem</span></span>|
|[<span data-ttu-id="8e966-806">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-806">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-807">新規作成</span><span class="sxs-lookup"><span data-stu-id="8e966-807">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8e966-808">例</span><span class="sxs-lookup"><span data-stu-id="8e966-808">Examples</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="8e966-809">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8e966-809">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="8e966-810">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="8e966-810">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="8e966-811">現在、サポートされているイベントの種類は `Office.EventType.AttachmentsChanged`、`Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged`、`Office.EventType.RecurrenceChanged` です。</span><span class="sxs-lookup"><span data-stu-id="8e966-811">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e966-812">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8e966-812">Parameters:</span></span>

| <span data-ttu-id="8e966-813">名前</span><span class="sxs-lookup"><span data-stu-id="8e966-813">Name</span></span> | <span data-ttu-id="8e966-814">型</span><span class="sxs-lookup"><span data-stu-id="8e966-814">Type</span></span> | <span data-ttu-id="8e966-815">属性</span><span class="sxs-lookup"><span data-stu-id="8e966-815">Attributes</span></span> | <span data-ttu-id="8e966-816">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-816">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="8e966-817">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="8e966-817">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="8e966-818">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="8e966-818">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="8e966-819">Function</span><span class="sxs-lookup"><span data-stu-id="8e966-819">Function</span></span> || <span data-ttu-id="8e966-p138">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="8e966-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="8e966-823">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-823">Object</span></span> | <span data-ttu-id="8e966-824">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-824">&lt;optional&gt;</span></span> | <span data-ttu-id="8e966-825">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8e966-825">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="8e966-826">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-826">Object</span></span> | <span data-ttu-id="8e966-827">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-827">&lt;optional&gt;</span></span> | <span data-ttu-id="8e966-828">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8e966-828">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="8e966-829">function</span><span class="sxs-lookup"><span data-stu-id="8e966-829">function</span></span>| <span data-ttu-id="8e966-830">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-830">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-831">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-831">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e966-832">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-832">Requirements</span></span>

|<span data-ttu-id="8e966-833">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-833">Requirement</span></span>| <span data-ttu-id="8e966-834">値</span><span class="sxs-lookup"><span data-stu-id="8e966-834">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-835">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-835">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e966-836">1.7</span><span class="sxs-lookup"><span data-stu-id="8e966-836">1.7</span></span> |
|[<span data-ttu-id="8e966-837">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-837">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e966-838">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-838">ReadItem</span></span> |
|[<span data-ttu-id="8e966-839">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-839">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e966-840">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="8e966-840">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="8e966-841">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8e966-841">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8e966-842">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="8e966-842">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="8e966-p139">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="8e966-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="8e966-846">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="8e966-846">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="8e966-847">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="8e966-847">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e966-848">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8e966-848">Parameters:</span></span>

|<span data-ttu-id="8e966-849">名前</span><span class="sxs-lookup"><span data-stu-id="8e966-849">Name</span></span>|<span data-ttu-id="8e966-850">型</span><span class="sxs-lookup"><span data-stu-id="8e966-850">Type</span></span>|<span data-ttu-id="8e966-851">属性</span><span class="sxs-lookup"><span data-stu-id="8e966-851">Attributes</span></span>|<span data-ttu-id="8e966-852">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-852">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="8e966-853">String</span><span class="sxs-lookup"><span data-stu-id="8e966-853">String</span></span>||<span data-ttu-id="8e966-p140">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="8e966-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="8e966-856">String</span><span class="sxs-lookup"><span data-stu-id="8e966-856">String</span></span>||<span data-ttu-id="8e966-p141">添付するアイテムの件名。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="8e966-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="8e966-859">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-859">Object</span></span>|<span data-ttu-id="8e966-860">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-860">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-861">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8e966-861">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8e966-862">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-862">Object</span></span>|<span data-ttu-id="8e966-863">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-863">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-864">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8e966-864">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8e966-865">function</span><span class="sxs-lookup"><span data-stu-id="8e966-865">function</span></span>|<span data-ttu-id="8e966-866">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-866">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-867">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-867">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8e966-868">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-868">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8e966-869">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="8e966-869">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8e966-870">エラー</span><span class="sxs-lookup"><span data-stu-id="8e966-870">Errors</span></span>

|<span data-ttu-id="8e966-871">エラー コード</span><span class="sxs-lookup"><span data-stu-id="8e966-871">Error code</span></span>|<span data-ttu-id="8e966-872">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-872">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="8e966-873">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="8e966-873">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e966-874">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-874">Requirements</span></span>

|<span data-ttu-id="8e966-875">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-875">Requirement</span></span>|<span data-ttu-id="8e966-876">値</span><span class="sxs-lookup"><span data-stu-id="8e966-876">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-877">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-877">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-878">1.1</span><span class="sxs-lookup"><span data-stu-id="8e966-878">1.1</span></span>|
|[<span data-ttu-id="8e966-879">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-879">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-880">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8e966-880">ReadWriteItem</span></span>|
|[<span data-ttu-id="8e966-881">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-881">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-882">作成</span><span class="sxs-lookup"><span data-stu-id="8e966-882">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-883">例</span><span class="sxs-lookup"><span data-stu-id="8e966-883">Example</span></span>

<span data-ttu-id="8e966-884">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-884">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="8e966-885">close()</span><span class="sxs-lookup"><span data-stu-id="8e966-885">close()</span></span>

<span data-ttu-id="8e966-886">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="8e966-886">Closes the current item that is being composed.</span></span>

<span data-ttu-id="8e966-p142">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="8e966-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="8e966-889">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-889">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="8e966-890">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="8e966-890">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-891">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-891">Requirements</span></span>

|<span data-ttu-id="8e966-892">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-892">Requirement</span></span>|<span data-ttu-id="8e966-893">値</span><span class="sxs-lookup"><span data-stu-id="8e966-893">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-894">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-894">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-895">1.3</span><span class="sxs-lookup"><span data-stu-id="8e966-895">1.3</span></span>|
|[<span data-ttu-id="8e966-896">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-896">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-897">制限あり</span><span class="sxs-lookup"><span data-stu-id="8e966-897">Restricted</span></span>|
|[<span data-ttu-id="8e966-898">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-898">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-899">作成</span><span class="sxs-lookup"><span data-stu-id="8e966-899">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="8e966-900">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="8e966-900">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="8e966-901">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-901">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8e966-902">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8e966-902">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8e966-903">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-903">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8e966-904">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="8e966-904">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="8e966-p143">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="8e966-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e966-908">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8e966-908">Parameters:</span></span>

|<span data-ttu-id="8e966-909">名前</span><span class="sxs-lookup"><span data-stu-id="8e966-909">Name</span></span>|<span data-ttu-id="8e966-910">型</span><span class="sxs-lookup"><span data-stu-id="8e966-910">Type</span></span>|<span data-ttu-id="8e966-911">属性</span><span class="sxs-lookup"><span data-stu-id="8e966-911">Attributes</span></span>|<span data-ttu-id="8e966-912">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-912">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="8e966-913">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="8e966-913">String &#124; Object</span></span>||<span data-ttu-id="8e966-p144">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="8e966-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8e966-916">**または**</span><span class="sxs-lookup"><span data-stu-id="8e966-916">**OR**</span></span><br/><span data-ttu-id="8e966-p145">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8e966-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="8e966-919">String</span><span class="sxs-lookup"><span data-stu-id="8e966-919">String</span></span>|<span data-ttu-id="8e966-920">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-920">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="8e966-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="8e966-923">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-923">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="8e966-924">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-924">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-925">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="8e966-925">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="8e966-926">String</span><span class="sxs-lookup"><span data-stu-id="8e966-926">String</span></span>||<span data-ttu-id="8e966-p147">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="8e966-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="8e966-929">String</span><span class="sxs-lookup"><span data-stu-id="8e966-929">String</span></span>||<span data-ttu-id="8e966-930">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="8e966-930">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="8e966-931">String</span><span class="sxs-lookup"><span data-stu-id="8e966-931">String</span></span>||<span data-ttu-id="8e966-p148">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="8e966-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="8e966-934">ブール値</span><span class="sxs-lookup"><span data-stu-id="8e966-934">Boolean</span></span>||<span data-ttu-id="8e966-p149">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="8e966-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="8e966-937">String</span><span class="sxs-lookup"><span data-stu-id="8e966-937">String</span></span>||<span data-ttu-id="8e966-p150">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="8e966-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="8e966-941">function</span><span class="sxs-lookup"><span data-stu-id="8e966-941">function</span></span>|<span data-ttu-id="8e966-942">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-942">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-943">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-943">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e966-944">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-944">Requirements</span></span>

|<span data-ttu-id="8e966-945">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-945">Requirement</span></span>|<span data-ttu-id="8e966-946">値</span><span class="sxs-lookup"><span data-stu-id="8e966-946">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-947">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-947">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-948">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-948">1.0</span></span>|
|[<span data-ttu-id="8e966-949">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-949">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-950">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-950">ReadItem</span></span>|
|[<span data-ttu-id="8e966-951">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-951">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-952">読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-952">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8e966-953">例</span><span class="sxs-lookup"><span data-stu-id="8e966-953">Examples</span></span>

<span data-ttu-id="8e966-954">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="8e966-954">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="8e966-955">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="8e966-955">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="8e966-956">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="8e966-956">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8e966-957">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="8e966-957">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="8e966-958">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="8e966-958">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="8e966-959">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="8e966-959">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="8e966-960">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="8e966-960">displayReplyForm(formData)</span></span>

<span data-ttu-id="8e966-961">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-961">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8e966-962">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8e966-962">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8e966-963">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-963">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8e966-964">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="8e966-964">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="8e966-p151">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="8e966-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e966-968">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8e966-968">Parameters:</span></span>

|<span data-ttu-id="8e966-969">名前</span><span class="sxs-lookup"><span data-stu-id="8e966-969">Name</span></span>|<span data-ttu-id="8e966-970">型</span><span class="sxs-lookup"><span data-stu-id="8e966-970">Type</span></span>|<span data-ttu-id="8e966-971">属性</span><span class="sxs-lookup"><span data-stu-id="8e966-971">Attributes</span></span>|<span data-ttu-id="8e966-972">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-972">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="8e966-973">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="8e966-973">String &#124; Object</span></span>||<span data-ttu-id="8e966-p152">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="8e966-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8e966-976">**または**</span><span class="sxs-lookup"><span data-stu-id="8e966-976">**OR**</span></span><br/><span data-ttu-id="8e966-p153">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8e966-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="8e966-979">String</span><span class="sxs-lookup"><span data-stu-id="8e966-979">String</span></span>|<span data-ttu-id="8e966-980">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-980">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-p154">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="8e966-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="8e966-983">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-983">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="8e966-984">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-984">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-985">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="8e966-985">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="8e966-986">String</span><span class="sxs-lookup"><span data-stu-id="8e966-986">String</span></span>||<span data-ttu-id="8e966-p155">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="8e966-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="8e966-989">String</span><span class="sxs-lookup"><span data-stu-id="8e966-989">String</span></span>||<span data-ttu-id="8e966-990">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="8e966-990">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="8e966-991">String</span><span class="sxs-lookup"><span data-stu-id="8e966-991">String</span></span>||<span data-ttu-id="8e966-p156">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="8e966-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="8e966-994">ブール値</span><span class="sxs-lookup"><span data-stu-id="8e966-994">Boolean</span></span>||<span data-ttu-id="8e966-p157">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="8e966-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="8e966-997">String</span><span class="sxs-lookup"><span data-stu-id="8e966-997">String</span></span>||<span data-ttu-id="8e966-p158">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="8e966-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="8e966-1001">function</span><span class="sxs-lookup"><span data-stu-id="8e966-1001">function</span></span>|<span data-ttu-id="8e966-1002">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1002">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1003">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1003">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e966-1004">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1004">Requirements</span></span>

|<span data-ttu-id="8e966-1005">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1005">Requirement</span></span>|<span data-ttu-id="8e966-1006">値</span><span class="sxs-lookup"><span data-stu-id="8e966-1006">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-1007">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-1007">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-1008">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-1008">1.0</span></span>|
|[<span data-ttu-id="8e966-1009">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-1009">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-1010">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-1010">ReadItem</span></span>|
|[<span data-ttu-id="8e966-1011">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-1011">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-1012">読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-1012">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8e966-1013">例</span><span class="sxs-lookup"><span data-stu-id="8e966-1013">Examples</span></span>

<span data-ttu-id="8e966-1014">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1014">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="8e966-1015">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1015">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="8e966-1016">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1016">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8e966-1017">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1017">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="8e966-1018">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1018">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="8e966-1019">本文、ファイルの添付ファイル、アイテムの添付ファイル、コールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1019">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="8e966-1020">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="8e966-1020">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="8e966-1021">メッセージまたは予定から指定の添付ファイルを取得し、それを `AttachmentContent` オブジェクトとして返されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1021">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="8e966-1022">`getAttachmentContentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから取得します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1022">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="8e966-1023">ベスト プラクティスとして、識別子を使用し、`getAttachmentsAsync` または `item.attachments` 呼び出しで attachmentIds を取得した同じセッションで添付ファイルを取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8e966-1023">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="8e966-1024">Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="8e966-1024">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="8e966-1025">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始し、フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1025">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e966-1026">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8e966-1026">Parameters:</span></span>

|<span data-ttu-id="8e966-1027">名前</span><span class="sxs-lookup"><span data-stu-id="8e966-1027">Name</span></span>|<span data-ttu-id="8e966-1028">型</span><span class="sxs-lookup"><span data-stu-id="8e966-1028">Type</span></span>|<span data-ttu-id="8e966-1029">属性</span><span class="sxs-lookup"><span data-stu-id="8e966-1029">Attributes</span></span>|<span data-ttu-id="8e966-1030">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-1030">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="8e966-1031">String</span><span class="sxs-lookup"><span data-stu-id="8e966-1031">String</span></span>||<span data-ttu-id="8e966-1032">取得する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="8e966-1032">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="8e966-1033">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-1033">Object</span></span>|<span data-ttu-id="8e966-1034">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1034">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1035">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8e966-1035">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8e966-1036">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-1036">Object</span></span>|<span data-ttu-id="8e966-1037">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1037">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1038">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1038">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8e966-1039">function</span><span class="sxs-lookup"><span data-stu-id="8e966-1039">function</span></span>|<span data-ttu-id="8e966-1040">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1040">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1041">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1041">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e966-1042">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1042">Requirements</span></span>

|<span data-ttu-id="8e966-1043">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1043">Requirement</span></span>|<span data-ttu-id="8e966-1044">値</span><span class="sxs-lookup"><span data-stu-id="8e966-1044">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-1045">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-1045">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-1046">プレビュー</span><span class="sxs-lookup"><span data-stu-id="8e966-1046">Preview</span></span>|
|[<span data-ttu-id="8e966-1047">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-1047">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-1048">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-1048">ReadItem</span></span>|
|[<span data-ttu-id="8e966-1049">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-1049">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-1050">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-1050">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e966-1051">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8e966-1051">Returns:</span></span>

<span data-ttu-id="8e966-1052">型: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="8e966-1052">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="8e966-1053">例</span><span class="sxs-lookup"><span data-stu-id="8e966-1053">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="8e966-1054">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="8e966-1054">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="8e966-1055">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1055">Gets the item's attachments as an array.</span></span> <span data-ttu-id="8e966-1056">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8e966-1056">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e966-1057">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8e966-1057">Parameters:</span></span>

|<span data-ttu-id="8e966-1058">名前</span><span class="sxs-lookup"><span data-stu-id="8e966-1058">Name</span></span>|<span data-ttu-id="8e966-1059">型</span><span class="sxs-lookup"><span data-stu-id="8e966-1059">Type</span></span>|<span data-ttu-id="8e966-1060">属性</span><span class="sxs-lookup"><span data-stu-id="8e966-1060">Attributes</span></span>|<span data-ttu-id="8e966-1061">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-1061">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="8e966-1062">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8e966-1062">Object</span></span>|<span data-ttu-id="8e966-1063">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1063">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1064">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8e966-1064">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8e966-1065">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-1065">Object</span></span>|<span data-ttu-id="8e966-1066">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1066">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1067">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1067">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8e966-1068">function</span><span class="sxs-lookup"><span data-stu-id="8e966-1068">function</span></span>|<span data-ttu-id="8e966-1069">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1070">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1070">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e966-1071">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1071">Requirements</span></span>

|<span data-ttu-id="8e966-1072">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1072">Requirement</span></span>|<span data-ttu-id="8e966-1073">値</span><span class="sxs-lookup"><span data-stu-id="8e966-1073">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-1074">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-1074">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-1075">プレビュー</span><span class="sxs-lookup"><span data-stu-id="8e966-1075">Preview</span></span>|
|[<span data-ttu-id="8e966-1076">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-1076">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-1077">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-1077">ReadItem</span></span>|
|[<span data-ttu-id="8e966-1078">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-1078">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-1079">作成</span><span class="sxs-lookup"><span data-stu-id="8e966-1079">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e966-1080">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8e966-1080">Returns:</span></span>

<span data-ttu-id="8e966-1081">型: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="8e966-1081">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="8e966-1082">例</span><span class="sxs-lookup"><span data-stu-id="8e966-1082">Example</span></span>

<span data-ttu-id="8e966-1083">次の例では、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1083">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="8e966-1084">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="8e966-1084">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="8e966-1085">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1085">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="8e966-1086">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8e966-1086">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-1087">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1087">Requirements</span></span>

|<span data-ttu-id="8e966-1088">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1088">Requirement</span></span>|<span data-ttu-id="8e966-1089">値</span><span class="sxs-lookup"><span data-stu-id="8e966-1089">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-1090">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-1090">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-1091">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-1091">1.0</span></span>|
|[<span data-ttu-id="8e966-1092">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-1092">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-1093">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-1093">ReadItem</span></span>|
|[<span data-ttu-id="8e966-1094">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-1094">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-1095">読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-1095">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e966-1096">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8e966-1096">Returns:</span></span>

<span data-ttu-id="8e966-1097">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="8e966-1097">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="8e966-1098">例</span><span class="sxs-lookup"><span data-stu-id="8e966-1098">Example</span></span>

<span data-ttu-id="8e966-1099">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="8e966-1099">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="8e966-1100">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="8e966-1100">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="8e966-1101">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1101">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="8e966-1102">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8e966-1102">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e966-1103">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8e966-1103">Parameters:</span></span>

|<span data-ttu-id="8e966-1104">名前</span><span class="sxs-lookup"><span data-stu-id="8e966-1104">Name</span></span>|<span data-ttu-id="8e966-1105">種類</span><span class="sxs-lookup"><span data-stu-id="8e966-1105">Type</span></span>|<span data-ttu-id="8e966-1106">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-1106">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="8e966-1107">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="8e966-1107">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="8e966-1108">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="8e966-1108">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e966-1109">Requirements</span><span class="sxs-lookup"><span data-stu-id="8e966-1109">Requirements</span></span>

|<span data-ttu-id="8e966-1110">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1110">Requirement</span></span>|<span data-ttu-id="8e966-1111">値</span><span class="sxs-lookup"><span data-stu-id="8e966-1111">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-1112">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-1112">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-1113">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-1113">1.0</span></span>|
|[<span data-ttu-id="8e966-1114">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-1114">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-1115">制限あり</span><span class="sxs-lookup"><span data-stu-id="8e966-1115">Restricted</span></span>|
|[<span data-ttu-id="8e966-1116">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-1116">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-1117">読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-1117">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e966-1118">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8e966-1118">Returns:</span></span>

<span data-ttu-id="8e966-1119">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1119">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="8e966-1120">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1120">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="8e966-1121">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="8e966-1121">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="8e966-1122">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="8e966-1122">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="8e966-1123">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="8e966-1123">Value of `entityType`</span></span>|<span data-ttu-id="8e966-1124">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="8e966-1124">Type of objects in returned array</span></span>|<span data-ttu-id="8e966-1125">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="8e966-1125">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="8e966-1126">String</span><span class="sxs-lookup"><span data-stu-id="8e966-1126">String</span></span>|<span data-ttu-id="8e966-1127">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="8e966-1127">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="8e966-1128">連絡先</span><span class="sxs-lookup"><span data-stu-id="8e966-1128">Contact</span></span>|<span data-ttu-id="8e966-1129">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8e966-1129">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="8e966-1130">文字列</span><span class="sxs-lookup"><span data-stu-id="8e966-1130">String</span></span>|<span data-ttu-id="8e966-1131">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8e966-1131">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="8e966-1132">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="8e966-1132">MeetingSuggestion</span></span>|<span data-ttu-id="8e966-1133">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8e966-1133">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="8e966-1134">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="8e966-1134">PhoneNumber</span></span>|<span data-ttu-id="8e966-1135">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="8e966-1135">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="8e966-1136">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="8e966-1136">TaskSuggestion</span></span>|<span data-ttu-id="8e966-1137">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8e966-1137">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="8e966-1138">String</span><span class="sxs-lookup"><span data-stu-id="8e966-1138">String</span></span>|<span data-ttu-id="8e966-1139">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="8e966-1139">**Restricted**</span></span>|

<span data-ttu-id="8e966-1140">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="8e966-1140">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="8e966-1141">例</span><span class="sxs-lookup"><span data-stu-id="8e966-1141">Example</span></span>

<span data-ttu-id="8e966-1142">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1142">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="8e966-1143">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="8e966-1143">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="8e966-1144">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1144">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8e966-1145">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8e966-1145">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8e966-1146">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1146">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e966-1147">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8e966-1147">Parameters:</span></span>

|<span data-ttu-id="8e966-1148">名前</span><span class="sxs-lookup"><span data-stu-id="8e966-1148">Name</span></span>|<span data-ttu-id="8e966-1149">種類</span><span class="sxs-lookup"><span data-stu-id="8e966-1149">Type</span></span>|<span data-ttu-id="8e966-1150">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-1150">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="8e966-1151">String</span><span class="sxs-lookup"><span data-stu-id="8e966-1151">String</span></span>|<span data-ttu-id="8e966-1152">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="8e966-1152">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e966-1153">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1153">Requirements</span></span>

|<span data-ttu-id="8e966-1154">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1154">Requirement</span></span>|<span data-ttu-id="8e966-1155">値</span><span class="sxs-lookup"><span data-stu-id="8e966-1155">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-1156">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-1156">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-1157">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-1157">1.0</span></span>|
|[<span data-ttu-id="8e966-1158">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-1158">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-1159">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-1159">ReadItem</span></span>|
|[<span data-ttu-id="8e966-1160">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-1160">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-1161">読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-1161">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e966-1162">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8e966-1162">Returns:</span></span>

<span data-ttu-id="8e966-p162">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-p162">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="8e966-1165">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="8e966-1165">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="8e966-1166">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8e966-1166">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="8e966-1167">アドインが[操作可能メッセージによってアクティブ化](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを取得します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1167">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="8e966-1168">このメソッドは、Outlook 2016 for Windows 以降 (16.0.8413.1000 以降のクイック実行バージョン) および Outlook on the web for Office 365 でのみサポートされます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1168">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e966-1169">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8e966-1169">Parameters:</span></span>
|<span data-ttu-id="8e966-1170">名前</span><span class="sxs-lookup"><span data-stu-id="8e966-1170">Name</span></span>|<span data-ttu-id="8e966-1171">型</span><span class="sxs-lookup"><span data-stu-id="8e966-1171">Type</span></span>|<span data-ttu-id="8e966-1172">属性</span><span class="sxs-lookup"><span data-stu-id="8e966-1172">Attributes</span></span>|<span data-ttu-id="8e966-1173">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-1173">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="8e966-1174">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8e966-1174">Object</span></span>|<span data-ttu-id="8e966-1175">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1175">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1176">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8e966-1176">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8e966-1177">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-1177">Object</span></span>|<span data-ttu-id="8e966-1178">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1178">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1179">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1179">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8e966-1180">function</span><span class="sxs-lookup"><span data-stu-id="8e966-1180">function</span></span>|<span data-ttu-id="8e966-1181">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1181">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1182">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1182">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8e966-1183">成功すると、初期化データが文字列として `asyncResult.value` プロパティで指定されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1183">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="8e966-1184">初期化コンテキストがない場合、`asyncResult` オブジェクトには、`code` プロパティが `9020`、`name` プロパティが `GenericResponseError` に設定された `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1184">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e966-1185">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1185">Requirements</span></span>

|<span data-ttu-id="8e966-1186">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1186">Requirement</span></span>|<span data-ttu-id="8e966-1187">値</span><span class="sxs-lookup"><span data-stu-id="8e966-1187">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-1188">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-1188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-1189">プレビュー</span><span class="sxs-lookup"><span data-stu-id="8e966-1189">Preview</span></span>|
|[<span data-ttu-id="8e966-1190">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-1190">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-1191">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-1191">ReadItem</span></span>|
|[<span data-ttu-id="8e966-1192">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-1192">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-1193">読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-1193">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-1194">例</span><span class="sxs-lookup"><span data-stu-id="8e966-1194">Example</span></span>

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

#### <a name="getregexmatches--object"></a><span data-ttu-id="8e966-1195">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="8e966-1195">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="8e966-1196">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1196">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8e966-1197">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8e966-1197">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8e966-p163">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="8e966-p163">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="8e966-1201">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="8e966-1201">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="8e966-1202">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="8e966-1202">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="8e966-p164">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="8e966-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-1206">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1206">Requirements</span></span>

|<span data-ttu-id="8e966-1207">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1207">Requirement</span></span>|<span data-ttu-id="8e966-1208">値</span><span class="sxs-lookup"><span data-stu-id="8e966-1208">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-1209">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-1209">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-1210">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-1210">1.0</span></span>|
|[<span data-ttu-id="8e966-1211">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-1211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-1212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-1212">ReadItem</span></span>|
|[<span data-ttu-id="8e966-1213">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-1213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-1214">読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-1214">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e966-1215">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8e966-1215">Returns:</span></span>

<span data-ttu-id="8e966-p165">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="8e966-p165">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="8e966-1218">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="8e966-1218">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8e966-1219">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-1219">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8e966-1220">例</span><span class="sxs-lookup"><span data-stu-id="8e966-1220">Example</span></span>

<span data-ttu-id="8e966-1221">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8e966-1221">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="8e966-1222">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="8e966-1222">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="8e966-1223">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1223">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8e966-1224">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8e966-1224">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8e966-1225">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="8e966-1225">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="8e966-p166">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="8e966-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e966-1228">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8e966-1228">Parameters:</span></span>

|<span data-ttu-id="8e966-1229">名前</span><span class="sxs-lookup"><span data-stu-id="8e966-1229">Name</span></span>|<span data-ttu-id="8e966-1230">種類</span><span class="sxs-lookup"><span data-stu-id="8e966-1230">Type</span></span>|<span data-ttu-id="8e966-1231">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-1231">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="8e966-1232">String</span><span class="sxs-lookup"><span data-stu-id="8e966-1232">String</span></span>|<span data-ttu-id="8e966-1233">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="8e966-1233">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e966-1234">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1234">Requirements</span></span>

|<span data-ttu-id="8e966-1235">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1235">Requirement</span></span>|<span data-ttu-id="8e966-1236">値</span><span class="sxs-lookup"><span data-stu-id="8e966-1236">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-1237">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-1237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-1238">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-1238">1.0</span></span>|
|[<span data-ttu-id="8e966-1239">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-1239">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-1240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-1240">ReadItem</span></span>|
|[<span data-ttu-id="8e966-1241">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-1241">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-1242">読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-1242">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e966-1243">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8e966-1243">Returns:</span></span>

<span data-ttu-id="8e966-1244">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="8e966-1244">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="8e966-1245">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="8e966-1245">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8e966-1246">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="8e966-1246">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8e966-1247">例</span><span class="sxs-lookup"><span data-stu-id="8e966-1247">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="8e966-1248">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="8e966-1248">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="8e966-1249">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1249">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="8e966-p167">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-p167">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e966-1252">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8e966-1252">Parameters:</span></span>

|<span data-ttu-id="8e966-1253">名前</span><span class="sxs-lookup"><span data-stu-id="8e966-1253">Name</span></span>|<span data-ttu-id="8e966-1254">型</span><span class="sxs-lookup"><span data-stu-id="8e966-1254">Type</span></span>|<span data-ttu-id="8e966-1255">属性</span><span class="sxs-lookup"><span data-stu-id="8e966-1255">Attributes</span></span>|<span data-ttu-id="8e966-1256">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-1256">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="8e966-1257">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8e966-1257">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="8e966-p168">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="8e966-p168">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="8e966-1261">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-1261">Object</span></span>|<span data-ttu-id="8e966-1262">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1262">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1263">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8e966-1263">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8e966-1264">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-1264">Object</span></span>|<span data-ttu-id="8e966-1265">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1265">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1266">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1266">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8e966-1267">function</span><span class="sxs-lookup"><span data-stu-id="8e966-1267">function</span></span>||<span data-ttu-id="8e966-1268">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1268">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8e966-1269">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1269">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="8e966-1270">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="8e966-1270">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e966-1271">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1271">Requirements</span></span>

|<span data-ttu-id="8e966-1272">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1272">Requirement</span></span>|<span data-ttu-id="8e966-1273">値</span><span class="sxs-lookup"><span data-stu-id="8e966-1273">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-1274">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-1274">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-1275">1.2</span><span class="sxs-lookup"><span data-stu-id="8e966-1275">1.2</span></span>|
|[<span data-ttu-id="8e966-1276">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-1276">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-1277">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8e966-1277">ReadWriteItem</span></span>|
|[<span data-ttu-id="8e966-1278">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-1278">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-1279">作成</span><span class="sxs-lookup"><span data-stu-id="8e966-1279">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e966-1280">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8e966-1280">Returns:</span></span>

<span data-ttu-id="8e966-1281">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="8e966-1281">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="8e966-1282">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="8e966-1282">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8e966-1283">String</span><span class="sxs-lookup"><span data-stu-id="8e966-1283">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8e966-1284">例</span><span class="sxs-lookup"><span data-stu-id="8e966-1284">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="8e966-1285">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="8e966-1285">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="8e966-p170">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-p170">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="8e966-1288">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8e966-1288">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-1289">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1289">Requirements</span></span>

|<span data-ttu-id="8e966-1290">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1290">Requirement</span></span>|<span data-ttu-id="8e966-1291">値</span><span class="sxs-lookup"><span data-stu-id="8e966-1291">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-1292">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-1292">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-1293">1.6</span><span class="sxs-lookup"><span data-stu-id="8e966-1293">1.6</span></span>|
|[<span data-ttu-id="8e966-1294">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-1294">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-1295">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-1295">ReadItem</span></span>|
|[<span data-ttu-id="8e966-1296">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-1296">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-1297">読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-1297">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e966-1298">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8e966-1298">Returns:</span></span>

<span data-ttu-id="8e966-1299">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="8e966-1299">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="8e966-1300">例</span><span class="sxs-lookup"><span data-stu-id="8e966-1300">Example</span></span>

<span data-ttu-id="8e966-1301">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="8e966-1301">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="8e966-1302">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="8e966-1302">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="8e966-p171">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-p171">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="8e966-1305">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8e966-1305">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8e966-p172">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="8e966-p172">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="8e966-1309">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="8e966-1309">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="8e966-1310">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="8e966-1310">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="8e966-p173">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="8e966-p173">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e966-1314">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1314">Requirements</span></span>

|<span data-ttu-id="8e966-1315">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1315">Requirement</span></span>|<span data-ttu-id="8e966-1316">値</span><span class="sxs-lookup"><span data-stu-id="8e966-1316">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-1317">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-1317">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-1318">1.6</span><span class="sxs-lookup"><span data-stu-id="8e966-1318">1.6</span></span>|
|[<span data-ttu-id="8e966-1319">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-1319">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-1320">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-1320">ReadItem</span></span>|
|[<span data-ttu-id="8e966-1321">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-1321">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-1322">読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-1322">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e966-1323">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8e966-1323">Returns:</span></span>

<span data-ttu-id="8e966-p174">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="8e966-p174">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="8e966-1326">例</span><span class="sxs-lookup"><span data-stu-id="8e966-1326">Example</span></span>

<span data-ttu-id="8e966-1327">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8e966-1327">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="8e966-1328">getSharedPropertiesAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="8e966-1328">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="8e966-1329">共有フォルダー、カレンダー、メールボックスで選択した予定またはメッセージのプロパティを取得します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1329">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e966-1330">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8e966-1330">Parameters:</span></span>

|<span data-ttu-id="8e966-1331">名前</span><span class="sxs-lookup"><span data-stu-id="8e966-1331">Name</span></span>|<span data-ttu-id="8e966-1332">型</span><span class="sxs-lookup"><span data-stu-id="8e966-1332">Type</span></span>|<span data-ttu-id="8e966-1333">属性</span><span class="sxs-lookup"><span data-stu-id="8e966-1333">Attributes</span></span>|<span data-ttu-id="8e966-1334">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-1334">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="8e966-1335">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-1335">Object</span></span>|<span data-ttu-id="8e966-1336">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1336">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1337">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8e966-1337">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8e966-1338">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-1338">Object</span></span>|<span data-ttu-id="8e966-1339">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1339">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1340">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1340">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8e966-1341">function</span><span class="sxs-lookup"><span data-stu-id="8e966-1341">function</span></span>||<span data-ttu-id="8e966-1342">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1342">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8e966-1343">共有プロパティは `asyncResult.value` プロパティの [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1343">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="8e966-1344">このオブジェクトは、アイテムの共有プロパティの取得に使用できます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1344">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e966-1345">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1345">Requirements</span></span>

|<span data-ttu-id="8e966-1346">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1346">Requirement</span></span>|<span data-ttu-id="8e966-1347">値</span><span class="sxs-lookup"><span data-stu-id="8e966-1347">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-1348">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-1348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-1349">プレビュー</span><span class="sxs-lookup"><span data-stu-id="8e966-1349">Preview</span></span>|
|[<span data-ttu-id="8e966-1350">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-1350">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-1351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-1351">ReadItem</span></span>|
|[<span data-ttu-id="8e966-1352">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-1352">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-1353">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-1353">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-1354">例</span><span class="sxs-lookup"><span data-stu-id="8e966-1354">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="8e966-1355">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8e966-1355">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="8e966-1356">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1356">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="8e966-p176">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="8e966-p176">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e966-1360">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8e966-1360">Parameters:</span></span>

|<span data-ttu-id="8e966-1361">名前</span><span class="sxs-lookup"><span data-stu-id="8e966-1361">Name</span></span>|<span data-ttu-id="8e966-1362">型</span><span class="sxs-lookup"><span data-stu-id="8e966-1362">Type</span></span>|<span data-ttu-id="8e966-1363">属性</span><span class="sxs-lookup"><span data-stu-id="8e966-1363">Attributes</span></span>|<span data-ttu-id="8e966-1364">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-1364">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="8e966-1365">function</span><span class="sxs-lookup"><span data-stu-id="8e966-1365">function</span></span>||<span data-ttu-id="8e966-1366">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1366">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8e966-1367">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1367">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="8e966-1368">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1368">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="8e966-1369">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-1369">Object</span></span>|<span data-ttu-id="8e966-1370">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1370">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1371">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1371">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="8e966-1372">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1372">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e966-1373">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1373">Requirements</span></span>

|<span data-ttu-id="8e966-1374">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1374">Requirement</span></span>|<span data-ttu-id="8e966-1375">値</span><span class="sxs-lookup"><span data-stu-id="8e966-1375">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-1376">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-1376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-1377">1.0</span><span class="sxs-lookup"><span data-stu-id="8e966-1377">1.0</span></span>|
|[<span data-ttu-id="8e966-1378">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-1378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-1379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-1379">ReadItem</span></span>|
|[<span data-ttu-id="8e966-1380">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-1380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-1381">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8e966-1381">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-1382">例</span><span class="sxs-lookup"><span data-stu-id="8e966-1382">Example</span></span>

<span data-ttu-id="8e966-p179">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="8e966-p179">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="8e966-1386">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8e966-1386">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="8e966-1387">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1387">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="8e966-1388">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1388">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="8e966-1389">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="8e966-1389">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="8e966-1390">Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="8e966-1390">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="8e966-1391">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始し、フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1391">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e966-1392">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8e966-1392">Parameters:</span></span>

|<span data-ttu-id="8e966-1393">名前</span><span class="sxs-lookup"><span data-stu-id="8e966-1393">Name</span></span>|<span data-ttu-id="8e966-1394">型</span><span class="sxs-lookup"><span data-stu-id="8e966-1394">Type</span></span>|<span data-ttu-id="8e966-1395">属性</span><span class="sxs-lookup"><span data-stu-id="8e966-1395">Attributes</span></span>|<span data-ttu-id="8e966-1396">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-1396">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="8e966-1397">String</span><span class="sxs-lookup"><span data-stu-id="8e966-1397">String</span></span>||<span data-ttu-id="8e966-1398">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="8e966-1398">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="8e966-1399">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8e966-1399">Object</span></span>|<span data-ttu-id="8e966-1400">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1400">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1401">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8e966-1401">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8e966-1402">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-1402">Object</span></span>|<span data-ttu-id="8e966-1403">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1403">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1404">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1404">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8e966-1405">function</span><span class="sxs-lookup"><span data-stu-id="8e966-1405">function</span></span>|<span data-ttu-id="8e966-1406">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1406">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1407">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1407">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8e966-1408">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1408">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8e966-1409">エラー</span><span class="sxs-lookup"><span data-stu-id="8e966-1409">Errors</span></span>

|<span data-ttu-id="8e966-1410">エラー コード</span><span class="sxs-lookup"><span data-stu-id="8e966-1410">Error code</span></span>|<span data-ttu-id="8e966-1411">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-1411">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="8e966-1412">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="8e966-1412">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e966-1413">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1413">Requirements</span></span>

|<span data-ttu-id="8e966-1414">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1414">Requirement</span></span>|<span data-ttu-id="8e966-1415">値</span><span class="sxs-lookup"><span data-stu-id="8e966-1415">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-1416">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-1416">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-1417">1.1</span><span class="sxs-lookup"><span data-stu-id="8e966-1417">1.1</span></span>|
|[<span data-ttu-id="8e966-1418">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-1418">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-1419">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8e966-1419">ReadWriteItem</span></span>|
|[<span data-ttu-id="8e966-1420">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-1420">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-1421">作成</span><span class="sxs-lookup"><span data-stu-id="8e966-1421">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-1422">例</span><span class="sxs-lookup"><span data-stu-id="8e966-1422">Example</span></span>

<span data-ttu-id="8e966-1423">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1423">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="8e966-1424">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8e966-1424">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="8e966-1425">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1425">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="8e966-1426">現在、サポートされているイベントの種類は `Office.EventType.AttachmentsChanged`、`Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged`、`Office.EventType.RecurrenceChanged` です。</span><span class="sxs-lookup"><span data-stu-id="8e966-1426">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e966-1427">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8e966-1427">Parameters:</span></span>

| <span data-ttu-id="8e966-1428">名前</span><span class="sxs-lookup"><span data-stu-id="8e966-1428">Name</span></span> | <span data-ttu-id="8e966-1429">型</span><span class="sxs-lookup"><span data-stu-id="8e966-1429">Type</span></span> | <span data-ttu-id="8e966-1430">属性</span><span class="sxs-lookup"><span data-stu-id="8e966-1430">Attributes</span></span> | <span data-ttu-id="8e966-1431">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-1431">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="8e966-1432">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="8e966-1432">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="8e966-1433">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="8e966-1433">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="8e966-1434">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-1434">Object</span></span> | <span data-ttu-id="8e966-1435">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1435">&lt;optional&gt;</span></span> | <span data-ttu-id="8e966-1436">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8e966-1436">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="8e966-1437">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-1437">Object</span></span> | <span data-ttu-id="8e966-1438">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1438">&lt;optional&gt;</span></span> | <span data-ttu-id="8e966-1439">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1439">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="8e966-1440">function</span><span class="sxs-lookup"><span data-stu-id="8e966-1440">function</span></span>| <span data-ttu-id="8e966-1441">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1441">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1442">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1442">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e966-1443">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1443">Requirements</span></span>

|<span data-ttu-id="8e966-1444">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1444">Requirement</span></span>| <span data-ttu-id="8e966-1445">値</span><span class="sxs-lookup"><span data-stu-id="8e966-1445">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-1446">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-1446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e966-1447">1.7</span><span class="sxs-lookup"><span data-stu-id="8e966-1447">1.7</span></span> |
|[<span data-ttu-id="8e966-1448">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-1448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e966-1449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e966-1449">ReadItem</span></span> |
|[<span data-ttu-id="8e966-1450">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-1450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e966-1451">Compose または Read</span><span class="sxs-lookup"><span data-stu-id="8e966-1451">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="8e966-1452">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="8e966-1452">saveAsync([options], callback)</span></span>

<span data-ttu-id="8e966-1453">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1453">Asynchronously saves an item.</span></span>

<span data-ttu-id="8e966-p181">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-p181">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="8e966-1457">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="8e966-1457">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="8e966-1458">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1458">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="8e966-p183">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="8e966-1462">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="8e966-1462">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="8e966-1463">Mac Outlook では、新規作成モードの会議で `saveAsync` をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="8e966-1463">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="8e966-1464">Mac Outlook では、会議で `saveAsync` を呼び出すとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1464">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="8e966-1465">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1465">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e966-1466">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8e966-1466">Parameters:</span></span>

|<span data-ttu-id="8e966-1467">名前</span><span class="sxs-lookup"><span data-stu-id="8e966-1467">Name</span></span>|<span data-ttu-id="8e966-1468">型</span><span class="sxs-lookup"><span data-stu-id="8e966-1468">Type</span></span>|<span data-ttu-id="8e966-1469">属性</span><span class="sxs-lookup"><span data-stu-id="8e966-1469">Attributes</span></span>|<span data-ttu-id="8e966-1470">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-1470">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="8e966-1471">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-1471">Object</span></span>|<span data-ttu-id="8e966-1472">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1472">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1473">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8e966-1473">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8e966-1474">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-1474">Object</span></span>|<span data-ttu-id="8e966-1475">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1475">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1476">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1476">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="8e966-1477">function</span><span class="sxs-lookup"><span data-stu-id="8e966-1477">function</span></span>||<span data-ttu-id="8e966-1478">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1478">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8e966-1479">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1479">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e966-1480">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1480">Requirements</span></span>

|<span data-ttu-id="8e966-1481">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1481">Requirement</span></span>|<span data-ttu-id="8e966-1482">値</span><span class="sxs-lookup"><span data-stu-id="8e966-1482">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-1483">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-1483">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-1484">1.3</span><span class="sxs-lookup"><span data-stu-id="8e966-1484">1.3</span></span>|
|[<span data-ttu-id="8e966-1485">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-1485">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-1486">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8e966-1486">ReadWriteItem</span></span>|
|[<span data-ttu-id="8e966-1487">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-1487">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-1488">新規作成</span><span class="sxs-lookup"><span data-stu-id="8e966-1488">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8e966-1489">例</span><span class="sxs-lookup"><span data-stu-id="8e966-1489">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="8e966-p185">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="8e966-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="8e966-1492">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="8e966-1492">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="8e966-1493">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="8e966-1493">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="8e966-p186">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="8e966-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e966-1497">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="8e966-1497">Parameters:</span></span>

|<span data-ttu-id="8e966-1498">名前</span><span class="sxs-lookup"><span data-stu-id="8e966-1498">Name</span></span>|<span data-ttu-id="8e966-1499">型</span><span class="sxs-lookup"><span data-stu-id="8e966-1499">Type</span></span>|<span data-ttu-id="8e966-1500">属性</span><span class="sxs-lookup"><span data-stu-id="8e966-1500">Attributes</span></span>|<span data-ttu-id="8e966-1501">説明</span><span class="sxs-lookup"><span data-stu-id="8e966-1501">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="8e966-1502">String</span><span class="sxs-lookup"><span data-stu-id="8e966-1502">String</span></span>||<span data-ttu-id="8e966-p187">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="8e966-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="8e966-1506">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-1506">Object</span></span>|<span data-ttu-id="8e966-1507">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1507">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1508">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8e966-1508">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="8e966-1509">Object</span><span class="sxs-lookup"><span data-stu-id="8e966-1509">Object</span></span>|<span data-ttu-id="8e966-1510">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1510">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-1511">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1511">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="8e966-1512">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8e966-1512">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="8e966-1513">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8e966-1513">&lt;optional&gt;</span></span>|<span data-ttu-id="8e966-p188">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-p188">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="8e966-p189">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-p189">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="8e966-1518">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1518">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="8e966-1519">function</span><span class="sxs-lookup"><span data-stu-id="8e966-1519">function</span></span>||<span data-ttu-id="8e966-1520">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8e966-1520">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e966-1521">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1521">Requirements</span></span>

|<span data-ttu-id="8e966-1522">要件</span><span class="sxs-lookup"><span data-stu-id="8e966-1522">Requirement</span></span>|<span data-ttu-id="8e966-1523">値</span><span class="sxs-lookup"><span data-stu-id="8e966-1523">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e966-1524">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8e966-1524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="8e966-1525">1.2</span><span class="sxs-lookup"><span data-stu-id="8e966-1525">1.2</span></span>|
|[<span data-ttu-id="8e966-1526">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8e966-1526">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="8e966-1527">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8e966-1527">ReadWriteItem</span></span>|
|[<span data-ttu-id="8e966-1528">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8e966-1528">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="8e966-1529">作成</span><span class="sxs-lookup"><span data-stu-id="8e966-1529">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8e966-1530">例</span><span class="sxs-lookup"><span data-stu-id="8e966-1530">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
