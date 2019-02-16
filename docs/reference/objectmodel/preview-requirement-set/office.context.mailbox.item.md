---
title: Office. アイテム-プレビュー要件セット
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: ed3926e7e77550f43b87306cf27cf1e96341bd82
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068246"
---
# <a name="item"></a><span data-ttu-id="c3a4a-102">item</span><span class="sxs-lookup"><span data-stu-id="c3a4a-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="c3a4a-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="c3a4a-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="c3a4a-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-106">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-106">Requirements</span></span>

|<span data-ttu-id="c3a4a-107">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-107">Requirement</span></span>|<span data-ttu-id="c3a4a-108">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-110">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-110">1.0</span></span>|
|[<span data-ttu-id="c3a4a-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="c3a4a-112">Restricted</span></span>|
|[<span data-ttu-id="c3a4a-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-114">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c3a4a-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-115">Members and methods</span></span>

| <span data-ttu-id="c3a4a-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-116">Member</span></span> | <span data-ttu-id="c3a4a-117">種類</span><span class="sxs-lookup"><span data-stu-id="c3a4a-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c3a4a-118">attachments</span><span class="sxs-lookup"><span data-stu-id="c3a4a-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="c3a4a-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-119">Member</span></span> |
| [<span data-ttu-id="c3a4a-120">bcc</span><span class="sxs-lookup"><span data-stu-id="c3a4a-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="c3a4a-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-121">Member</span></span> |
| [<span data-ttu-id="c3a4a-122">body</span><span class="sxs-lookup"><span data-stu-id="c3a4a-122">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="c3a4a-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-123">Member</span></span> |
| [<span data-ttu-id="c3a4a-124">cc</span><span class="sxs-lookup"><span data-stu-id="c3a4a-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="c3a4a-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-125">Member</span></span> |
| [<span data-ttu-id="c3a4a-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="c3a4a-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="c3a4a-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-127">Member</span></span> |
| [<span data-ttu-id="c3a4a-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="c3a4a-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="c3a4a-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-129">Member</span></span> |
| [<span data-ttu-id="c3a4a-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="c3a4a-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="c3a4a-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-131">Member</span></span> |
| [<span data-ttu-id="c3a4a-132">end</span><span class="sxs-lookup"><span data-stu-id="c3a4a-132">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="c3a4a-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-133">Member</span></span> |
| [<span data-ttu-id="c3a4a-134">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="c3a4a-134">enhancedLocation</span></span>](#enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation) | <span data-ttu-id="c3a4a-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-135">Member</span></span> |
| [<span data-ttu-id="c3a4a-136">from</span><span class="sxs-lookup"><span data-stu-id="c3a4a-136">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="c3a4a-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-137">Member</span></span> |
| [<span data-ttu-id="c3a4a-138">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="c3a4a-138">internetHeaders</span></span>](#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders) | <span data-ttu-id="c3a4a-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-139">Member</span></span> |
| [<span data-ttu-id="c3a4a-140">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="c3a4a-140">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="c3a4a-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-141">Member</span></span> |
| [<span data-ttu-id="c3a4a-142">itemClass</span><span class="sxs-lookup"><span data-stu-id="c3a4a-142">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="c3a4a-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-143">Member</span></span> |
| [<span data-ttu-id="c3a4a-144">itemId</span><span class="sxs-lookup"><span data-stu-id="c3a4a-144">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="c3a4a-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-145">Member</span></span> |
| [<span data-ttu-id="c3a4a-146">itemType</span><span class="sxs-lookup"><span data-stu-id="c3a4a-146">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="c3a4a-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-147">Member</span></span> |
| [<span data-ttu-id="c3a4a-148">location</span><span class="sxs-lookup"><span data-stu-id="c3a4a-148">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="c3a4a-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-149">Member</span></span> |
| [<span data-ttu-id="c3a4a-150">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="c3a4a-150">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="c3a4a-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-151">Member</span></span> |
| [<span data-ttu-id="c3a4a-152">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="c3a4a-152">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="c3a4a-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-153">Member</span></span> |
| [<span data-ttu-id="c3a4a-154">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="c3a4a-154">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="c3a4a-155">Member</span><span class="sxs-lookup"><span data-stu-id="c3a4a-155">Member</span></span> |
| [<span data-ttu-id="c3a4a-156">organizer</span><span class="sxs-lookup"><span data-stu-id="c3a4a-156">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="c3a4a-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-157">Member</span></span> |
| [<span data-ttu-id="c3a4a-158">recurrence</span><span class="sxs-lookup"><span data-stu-id="c3a4a-158">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="c3a4a-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-159">Member</span></span> |
| [<span data-ttu-id="c3a4a-160">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="c3a4a-160">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="c3a4a-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-161">Member</span></span> |
| [<span data-ttu-id="c3a4a-162">sender</span><span class="sxs-lookup"><span data-stu-id="c3a4a-162">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="c3a4a-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-163">Member</span></span> |
| [<span data-ttu-id="c3a4a-164">seriesId</span><span class="sxs-lookup"><span data-stu-id="c3a4a-164">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="c3a4a-165">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-165">Member</span></span> |
| [<span data-ttu-id="c3a4a-166">start</span><span class="sxs-lookup"><span data-stu-id="c3a4a-166">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="c3a4a-167">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-167">Member</span></span> |
| [<span data-ttu-id="c3a4a-168">subject</span><span class="sxs-lookup"><span data-stu-id="c3a4a-168">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="c3a4a-169">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-169">Member</span></span> |
| [<span data-ttu-id="c3a4a-170">to</span><span class="sxs-lookup"><span data-stu-id="c3a4a-170">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="c3a4a-171">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-171">Member</span></span> |
| [<span data-ttu-id="c3a4a-172">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c3a4a-172">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="c3a4a-173">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-173">Method</span></span> |
| [<span data-ttu-id="c3a4a-174">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="c3a4a-174">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="c3a4a-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-175">Method</span></span> |
| [<span data-ttu-id="c3a4a-176">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c3a4a-176">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="c3a4a-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-177">Method</span></span> |
| [<span data-ttu-id="c3a4a-178">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c3a4a-178">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="c3a4a-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-179">Method</span></span> |
| [<span data-ttu-id="c3a4a-180">close</span><span class="sxs-lookup"><span data-stu-id="c3a4a-180">close</span></span>](#close) | <span data-ttu-id="c3a4a-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-181">Method</span></span> |
| [<span data-ttu-id="c3a4a-182">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="c3a4a-182">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="c3a4a-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-183">Method</span></span> |
| [<span data-ttu-id="c3a4a-184">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="c3a4a-184">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="c3a4a-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-185">Method</span></span> |
| [<span data-ttu-id="c3a4a-186">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="c3a4a-186">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) | <span data-ttu-id="c3a4a-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-187">Method</span></span> |
| [<span data-ttu-id="c3a4a-188">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="c3a4a-188">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="c3a4a-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-189">Method</span></span> |
| [<span data-ttu-id="c3a4a-190">getEntities</span><span class="sxs-lookup"><span data-stu-id="c3a4a-190">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="c3a4a-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-191">Method</span></span> |
| [<span data-ttu-id="c3a4a-192">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="c3a4a-192">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="c3a4a-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-193">Method</span></span> |
| [<span data-ttu-id="c3a4a-194">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="c3a4a-194">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="c3a4a-195">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-195">Method</span></span> |
| [<span data-ttu-id="c3a4a-196">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="c3a4a-196">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="c3a4a-197">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-197">Method</span></span> |
| [<span data-ttu-id="c3a4a-198">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="c3a4a-198">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="c3a4a-199">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-199">Method</span></span> |
| [<span data-ttu-id="c3a4a-200">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="c3a4a-200">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="c3a4a-201">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-201">Method</span></span> |
| [<span data-ttu-id="c3a4a-202">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c3a4a-202">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="c3a4a-203">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-203">Method</span></span> |
| [<span data-ttu-id="c3a4a-204">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="c3a4a-204">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="c3a4a-205">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-205">Method</span></span> |
| [<span data-ttu-id="c3a4a-206">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="c3a4a-206">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="c3a4a-207">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-207">Method</span></span> |
| [<span data-ttu-id="c3a4a-208">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="c3a4a-208">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="c3a4a-209">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-209">Method</span></span> |
| [<span data-ttu-id="c3a4a-210">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="c3a4a-210">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="c3a4a-211">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-211">Method</span></span> |
| [<span data-ttu-id="c3a4a-212">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c3a4a-212">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="c3a4a-213">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-213">Method</span></span> |
| [<span data-ttu-id="c3a4a-214">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c3a4a-214">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="c3a4a-215">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-215">Method</span></span> |
| [<span data-ttu-id="c3a4a-216">saveAsync</span><span class="sxs-lookup"><span data-stu-id="c3a4a-216">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="c3a4a-217">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-217">Method</span></span> |
| [<span data-ttu-id="c3a4a-218">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c3a4a-218">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="c3a4a-219">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-219">Method</span></span> |

### <a name="example"></a><span data-ttu-id="c3a4a-220">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-220">Example</span></span>

<span data-ttu-id="c3a4a-221">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-221">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="c3a4a-222">メンバー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-222">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="c3a4a-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c3a4a-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="c3a4a-224">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-224">Gets the item's attachments as an array.</span></span> <span data-ttu-id="c3a4a-225">閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-225">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a4a-226">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-226">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="c3a4a-227">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-227">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="c3a4a-228">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-228">Type</span></span>

*   <span data-ttu-id="c3a4a-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c3a4a-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-230">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-230">Requirements</span></span>

|<span data-ttu-id="c3a4a-231">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-231">Requirement</span></span>|<span data-ttu-id="c3a4a-232">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-233">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-234">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-234">1.0</span></span>|
|[<span data-ttu-id="c3a4a-235">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-235">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-236">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-236">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-237">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-237">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-238">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-238">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-239">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-239">Example</span></span>

<span data-ttu-id="c3a4a-240">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-240">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="c3a4a-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="c3a4a-242">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-242">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="c3a4a-243">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-243">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a4a-244">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-244">Type</span></span>

*   [<span data-ttu-id="c3a4a-245">Recipients</span><span class="sxs-lookup"><span data-stu-id="c3a4a-245">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="c3a4a-246">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-246">Requirements</span></span>

|<span data-ttu-id="c3a4a-247">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-247">Requirement</span></span>|<span data-ttu-id="c3a4a-248">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-249">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-250">1.1</span><span class="sxs-lookup"><span data-stu-id="c3a4a-250">1.1</span></span>|
|[<span data-ttu-id="c3a4a-251">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-251">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-252">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-253">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-253">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-254">作成</span><span class="sxs-lookup"><span data-stu-id="c3a4a-254">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-255">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-255">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="c3a4a-256">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-256">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="c3a4a-257">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-257">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a4a-258">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-258">Type</span></span>

*   [<span data-ttu-id="c3a4a-259">Body</span><span class="sxs-lookup"><span data-stu-id="c3a4a-259">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="c3a4a-260">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-260">Requirements</span></span>

|<span data-ttu-id="c3a4a-261">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-261">Requirement</span></span>|<span data-ttu-id="c3a4a-262">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-263">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-264">1.1</span><span class="sxs-lookup"><span data-stu-id="c3a4a-264">1.1</span></span>|
|[<span data-ttu-id="c3a4a-265">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-265">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-266">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-267">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-267">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-268">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-268">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-269">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-269">Example</span></span>

<span data-ttu-id="c3a4a-270">この例では、メッセージの本文をプレーンテキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-270">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="c3a4a-271">次の例は、コールバック関数に渡される result パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-271">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="c3a4a-272">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-272">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="c3a4a-273">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-273">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="c3a4a-274">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-274">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a4a-275">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-275">Read mode</span></span>

<span data-ttu-id="c3a4a-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a4a-278">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-278">Compose mode</span></span>

<span data-ttu-id="c3a4a-279">`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-279">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c3a4a-280">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-280">Type</span></span>

*   <span data-ttu-id="c3a4a-281">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-281">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-282">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-282">Requirements</span></span>

|<span data-ttu-id="c3a4a-283">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-283">Requirement</span></span>|<span data-ttu-id="c3a4a-284">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-285">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-285">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-286">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-286">1.0</span></span>|
|[<span data-ttu-id="c3a4a-287">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-287">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-288">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-289">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-289">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-290">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-290">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="c3a4a-291">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-291">(nullable) conversationId :String</span></span>

<span data-ttu-id="c3a4a-292">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-292">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="c3a4a-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="c3a4a-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a4a-297">Type</span><span class="sxs-lookup"><span data-stu-id="c3a4a-297">Type</span></span>

*   <span data-ttu-id="c3a4a-298">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-298">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-299">Requirements</span><span class="sxs-lookup"><span data-stu-id="c3a4a-299">Requirements</span></span>

|<span data-ttu-id="c3a4a-300">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-300">Requirement</span></span>|<span data-ttu-id="c3a4a-301">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-302">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-303">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-303">1.0</span></span>|
|[<span data-ttu-id="c3a4a-304">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-304">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-305">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-306">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-306">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-307">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-307">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-308">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-308">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="c3a4a-309">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="c3a4a-309">dateTimeCreated :Date</span></span>

<span data-ttu-id="c3a4a-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a4a-312">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-312">Type</span></span>

*   <span data-ttu-id="c3a4a-313">日付</span><span class="sxs-lookup"><span data-stu-id="c3a4a-313">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-314">Requirements</span><span class="sxs-lookup"><span data-stu-id="c3a4a-314">Requirements</span></span>

|<span data-ttu-id="c3a4a-315">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-315">Requirement</span></span>|<span data-ttu-id="c3a4a-316">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-316">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-317">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-317">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-318">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-318">1.0</span></span>|
|[<span data-ttu-id="c3a4a-319">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-319">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-320">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-320">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-321">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-321">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-322">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-322">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-323">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-323">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="c3a4a-324">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="c3a4a-324">dateTimeModified :Date</span></span>

<span data-ttu-id="c3a4a-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a4a-327">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-327">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a4a-328">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-328">Type</span></span>

*   <span data-ttu-id="c3a4a-329">日付</span><span class="sxs-lookup"><span data-stu-id="c3a4a-329">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-330">Requirements</span><span class="sxs-lookup"><span data-stu-id="c3a4a-330">Requirements</span></span>

|<span data-ttu-id="c3a4a-331">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-331">Requirement</span></span>|<span data-ttu-id="c3a4a-332">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-333">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-334">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-334">1.0</span></span>|
|[<span data-ttu-id="c3a4a-335">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-335">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-336">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-337">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-337">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-338">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-338">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-339">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-339">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="c3a4a-340">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-340">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="c3a4a-341">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-341">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="c3a4a-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a4a-344">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-344">Read mode</span></span>

<span data-ttu-id="c3a4a-345">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-345">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a4a-346">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-346">Compose mode</span></span>

<span data-ttu-id="c3a4a-347">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-347">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="c3a4a-348">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-348">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="c3a4a-349">次の例では、 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) `Time`オブジェクトのメソッドを使用して予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-349">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="c3a4a-350">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-350">Type</span></span>

*   <span data-ttu-id="c3a4a-351">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-351">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-352">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-352">Requirements</span></span>

|<span data-ttu-id="c3a4a-353">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-353">Requirement</span></span>|<span data-ttu-id="c3a4a-354">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-354">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-355">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-356">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-356">1.0</span></span>|
|[<span data-ttu-id="c3a4a-357">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-357">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-358">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-359">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-359">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-360">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-360">Compose or Read</span></span>|

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="c3a4a-361">enhancedLocation:[enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-361">enhancedLocation :[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="c3a4a-362">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-362">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a4a-363">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-363">Read mode</span></span>

<span data-ttu-id="c3a4a-364">この`enhancedLocation`プロパティは、予定に関連付けられている場所 ( [locationdetails](/javascript/api/outlook/office.locationdetails)オブジェクトで表される) のセットを取得できる[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-364">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c3a4a-365">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-365">Compose mode</span></span>

<span data-ttu-id="c3a4a-366">この`enhancedLocation`プロパティは、予定の場所を取得、削除、または追加するためのメソッドを提供する[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-366">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a4a-367">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-367">Type</span></span>

*   [<span data-ttu-id="c3a4a-368">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="c3a4a-368">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="c3a4a-369">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-369">Requirements</span></span>

|<span data-ttu-id="c3a4a-370">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-370">Requirement</span></span>|<span data-ttu-id="c3a4a-371">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-371">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-372">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-372">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-373">プレビュー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-373">Preview</span></span>|
|[<span data-ttu-id="c3a4a-374">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-374">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-375">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-375">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-376">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-376">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-377">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-377">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-378">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-378">Example</span></span>

<span data-ttu-id="c3a4a-379">次の例では、予定に関連付けられている現在の場所を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-379">The following example gets the current locations associated with the appointment.</span></span>

```javascript
Office.context.mailbox.item.enhancedLocation.getAsync(callbackFunction);

function callbackFunction(asyncResult) {
  asyncResult.value.forEach(function (place) {
    console.log("Display name: " + place.displayName);
    console.log("Type: " + place.locationIdentifier.type);
    if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
      console.log("Email address: " + place.emailAddress);
    }
  });
}
```

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="c3a4a-380">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-380">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="c3a4a-381">メッセージの送信者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-381">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="c3a4a-p112">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a4a-384">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-384">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a4a-385">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-385">Read mode</span></span>

<span data-ttu-id="c3a4a-386">`from` プロパティは `EmailAddressDetails` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-386">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a4a-387">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-387">Compose mode</span></span>

<span data-ttu-id="c3a4a-388">`from` プロパティは From 値を取得するメソッドを提供する `From` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-388">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c3a4a-389">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-389">Type</span></span>

*   <span data-ttu-id="c3a4a-390">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-390">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-391">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-391">Requirements</span></span>

|<span data-ttu-id="c3a4a-392">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-392">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="c3a4a-393">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-393">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-394">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-394">1.0</span></span>|<span data-ttu-id="c3a4a-395">1.7</span><span class="sxs-lookup"><span data-stu-id="c3a4a-395">1.7</span></span>|
|[<span data-ttu-id="c3a4a-396">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-397">ReadItem</span></span>|<span data-ttu-id="c3a4a-398">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-398">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3a4a-399">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-400">Read</span><span class="sxs-lookup"><span data-stu-id="c3a4a-400">Read</span></span>|<span data-ttu-id="c3a4a-401">Compose</span><span class="sxs-lookup"><span data-stu-id="c3a4a-401">Compose</span></span>|

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="c3a4a-402">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-402">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="c3a4a-403">メッセージのインターネット ヘッダーを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-403">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a4a-404">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-404">Type</span></span>

*   [<span data-ttu-id="c3a4a-405">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="c3a4a-405">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="c3a4a-406">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-406">Requirements</span></span>

|<span data-ttu-id="c3a4a-407">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-407">Requirement</span></span>|<span data-ttu-id="c3a4a-408">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-409">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-410">プレビュー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-410">Preview</span></span>|
|[<span data-ttu-id="c3a4a-411">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-412">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-413">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-414">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-414">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-415">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-415">Example</span></span>

```javascript
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="c3a4a-416">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-416">internetMessageId :String</span></span>

<span data-ttu-id="c3a4a-p113">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a4a-419">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-419">Type</span></span>

*   <span data-ttu-id="c3a4a-420">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-420">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-421">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-421">Requirements</span></span>

|<span data-ttu-id="c3a4a-422">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-422">Requirement</span></span>|<span data-ttu-id="c3a4a-423">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-424">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-425">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-425">1.0</span></span>|
|[<span data-ttu-id="c3a4a-426">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-427">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-428">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-429">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-430">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-430">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

#### <a name="itemclass-string"></a><span data-ttu-id="c3a4a-431">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-431">itemClass :String</span></span>

<span data-ttu-id="c3a4a-p114">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="c3a4a-p115">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="c3a4a-436">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-436">Type</span></span>|<span data-ttu-id="c3a4a-437">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-437">Description</span></span>|<span data-ttu-id="c3a4a-438">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="c3a4a-438">item class</span></span>|
|---|---|---|
|<span data-ttu-id="c3a4a-439">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="c3a4a-439">Appointment items</span></span>|<span data-ttu-id="c3a4a-440">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-440">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="c3a4a-441">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="c3a4a-441">Message items</span></span>|<span data-ttu-id="c3a4a-442">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-442">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="c3a4a-443">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-443">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a4a-444">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-444">Type</span></span>

*   <span data-ttu-id="c3a4a-445">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-445">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-446">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-446">Requirements</span></span>

|<span data-ttu-id="c3a4a-447">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-447">Requirement</span></span>|<span data-ttu-id="c3a4a-448">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-448">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-449">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-449">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-450">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-450">1.0</span></span>|
|[<span data-ttu-id="c3a4a-451">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-451">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-452">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-452">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-453">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-453">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-454">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-454">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-455">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-455">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="c3a4a-456">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-456">(nullable) itemId :String</span></span>

<span data-ttu-id="c3a4a-p116">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a4a-459">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-459">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="c3a4a-460">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-460">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="c3a4a-461">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-461">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="c3a4a-462">詳細は、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-462">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="c3a4a-p118">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a4a-465">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-465">Type</span></span>

*   <span data-ttu-id="c3a4a-466">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-466">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-467">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-467">Requirements</span></span>

|<span data-ttu-id="c3a4a-468">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-468">Requirement</span></span>|<span data-ttu-id="c3a4a-469">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-470">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-471">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-471">1.0</span></span>|
|[<span data-ttu-id="c3a4a-472">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-472">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-473">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-474">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-474">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-475">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-475">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-476">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-476">Example</span></span>

<span data-ttu-id="c3a4a-p119">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="c3a4a-479">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-479">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="c3a4a-480">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-480">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="c3a4a-481">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-481">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a4a-482">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-482">Type</span></span>

*   [<span data-ttu-id="c3a4a-483">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="c3a4a-483">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="c3a4a-484">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-484">Requirements</span></span>

|<span data-ttu-id="c3a4a-485">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-485">Requirement</span></span>|<span data-ttu-id="c3a4a-486">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-487">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-487">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-488">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-488">1.0</span></span>|
|[<span data-ttu-id="c3a4a-489">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-489">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-490">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-491">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-491">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-492">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-492">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-493">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-493">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="c3a4a-494">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-494">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="c3a4a-495">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-495">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a4a-496">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-496">Read mode</span></span>

<span data-ttu-id="c3a4a-497">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-497">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a4a-498">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-498">Compose mode</span></span>

<span data-ttu-id="c3a4a-499">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-499">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c3a4a-500">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-500">Type</span></span>

*   <span data-ttu-id="c3a4a-501">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-501">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-502">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-502">Requirements</span></span>

|<span data-ttu-id="c3a4a-503">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-503">Requirement</span></span>|<span data-ttu-id="c3a4a-504">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-505">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-506">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-506">1.0</span></span>|
|[<span data-ttu-id="c3a4a-507">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-507">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-508">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-509">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-509">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-510">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-510">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="c3a4a-511">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-511">normalizedSubject :String</span></span>

<span data-ttu-id="c3a4a-p120">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="c3a4a-p121">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a4a-516">Type</span><span class="sxs-lookup"><span data-stu-id="c3a4a-516">Type</span></span>

*   <span data-ttu-id="c3a4a-517">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-517">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-518">Requirements</span><span class="sxs-lookup"><span data-stu-id="c3a4a-518">Requirements</span></span>

|<span data-ttu-id="c3a4a-519">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-519">Requirement</span></span>|<span data-ttu-id="c3a4a-520">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-521">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-522">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-522">1.0</span></span>|
|[<span data-ttu-id="c3a4a-523">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-524">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-525">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-526">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-526">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-527">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-527">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="c3a4a-528">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-528">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="c3a4a-529">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-529">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a4a-530">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-530">Type</span></span>

*   [<span data-ttu-id="c3a4a-531">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="c3a4a-531">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="c3a4a-532">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-532">Requirements</span></span>

|<span data-ttu-id="c3a4a-533">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-533">Requirement</span></span>|<span data-ttu-id="c3a4a-534">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-534">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-535">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-535">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-536">1.3</span><span class="sxs-lookup"><span data-stu-id="c3a4a-536">1.3</span></span>|
|[<span data-ttu-id="c3a4a-537">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-537">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-538">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-538">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-539">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-539">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-540">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-540">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-541">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-541">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="c3a4a-542">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-542">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="c3a4a-543">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-543">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="c3a4a-544">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-544">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a4a-545">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-545">Read mode</span></span>

<span data-ttu-id="c3a4a-546">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-546">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a4a-547">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-547">Compose mode</span></span>

<span data-ttu-id="c3a4a-548">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-548">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c3a4a-549">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-549">Type</span></span>

*   <span data-ttu-id="c3a4a-550">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-550">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-551">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-551">Requirements</span></span>

|<span data-ttu-id="c3a4a-552">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-552">Requirement</span></span>|<span data-ttu-id="c3a4a-553">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-553">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-554">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-554">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-555">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-555">1.0</span></span>|
|[<span data-ttu-id="c3a4a-556">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-556">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-557">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-557">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-558">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-558">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-559">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-559">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="c3a4a-560">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-560">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="c3a4a-561">指定の会議の開催者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-561">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a4a-562">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-562">Read mode</span></span>

<span data-ttu-id="c3a4a-563">`organizer` プロパティは、会議開催者を表す [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-563">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a4a-564">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-564">Compose mode</span></span>

<span data-ttu-id="c3a4a-565">`organizer` プロパティは Organizer 値を取得するメソッドを提供する [Organizer](/javascript/api/outlook/office.organizer) オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-565">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="c3a4a-566">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-566">Type</span></span>

*   <span data-ttu-id="c3a4a-567">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-567">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-568">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-568">Requirements</span></span>

|<span data-ttu-id="c3a4a-569">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-569">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="c3a4a-570">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-571">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-571">1.0</span></span>|<span data-ttu-id="c3a4a-572">1.7</span><span class="sxs-lookup"><span data-stu-id="c3a4a-572">1.7</span></span>|
|[<span data-ttu-id="c3a4a-573">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-573">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-574">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-574">ReadItem</span></span>|<span data-ttu-id="c3a4a-575">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-575">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3a4a-576">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-576">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-577">Read</span><span class="sxs-lookup"><span data-stu-id="c3a4a-577">Read</span></span>|<span data-ttu-id="c3a4a-578">Compose</span><span class="sxs-lookup"><span data-stu-id="c3a4a-578">Compose</span></span>|

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="c3a4a-579">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-579">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="c3a4a-580">予定の定期的なパターンを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-580">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="c3a4a-581">会議出席依頼の定期的なパターンを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-581">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="c3a4a-582">予定アイテムの閲覧モードと新規作成モード。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-582">Read and compose modes for appointment items.</span></span> <span data-ttu-id="c3a4a-583">会議出席依頼アイテムの閲覧モード。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-583">Read mode for meeting request items.</span></span>

<span data-ttu-id="c3a4a-584">`recurrence` プロパティは、アイテムがシリーズか、シリーズに含まれるインスタンスの場合、定期的な予定または会議出席依頼に対して [recurrence](/javascript/api/outlook/office.recurrence) オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-584">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="c3a4a-585">`null` は、単発の予定および単発の予定の会議出席依頼に対して返されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-585">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="c3a4a-586">`undefined` は、会議出席依頼ではないメッセージに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-586">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="c3a4a-587">注: 会議出席依頼の `itemClass` 値は IPM.Schedule.Meeting.Request です。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-587">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="c3a4a-588">注: recurrence オブジェクトが `null` の場合、オブジェクトがシリーズの一部ではなく、1 つの単発の予定または 1 つの単発の予定の会議出席依頼であることを示します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-588">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a4a-589">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-589">Read mode</span></span>

<span data-ttu-id="c3a4a-590">この`recurrence`プロパティは、定期的な予定を表す[定期的](/javascript/api/outlook/office.recurrence)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-590">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="c3a4a-591">これは、予定および会議出席依頼に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-591">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a4a-592">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-592">Compose mode</span></span>

<span data-ttu-id="c3a4a-593">この`recurrence`プロパティは、予定の繰り返しを管理するためのメソッドを提供する[定期的](/javascript/api/outlook/office.recurrence)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-593">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="c3a4a-594">これは予定に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-594">This is available for appointments.</span></span>

```javascript
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var recurrence = asyncResult.value;
  if (!recurrence) {
    console.log("One-time appointment or meeting");
  } else {
    console.log(JSON.stringify(recurrence));
  }
}

// The following example shows the results of the getAsync call that retrieves the recurrence for a series.
// NOTE: In this example, seriesTimeObject is a placeholder for the JSON representing the
// recurrence.seriesTime property. You should use the SeriesTime object's methods to get the
// recurrence date and time properties.
Recurrence = {
  "recurrenceType": "weekly",
  "recurrenceProperties": {"interval": 2, "days": ["mon","thu","fri"], "firstDayOfWeek": "sun"},
  "seriesTime": {seriesTimeObject},
  "recurrenceTimeZone": {"name": "Pacific Standard Time", "offset": -480}
}
```

##### <a name="type"></a><span data-ttu-id="c3a4a-595">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-595">Type</span></span>

* [<span data-ttu-id="c3a4a-596">Recurrence</span><span class="sxs-lookup"><span data-stu-id="c3a4a-596">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="c3a4a-597">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-597">Requirement</span></span>|<span data-ttu-id="c3a4a-598">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-598">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-599">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-599">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-600">1.7</span><span class="sxs-lookup"><span data-stu-id="c3a4a-600">1.7</span></span>|
|[<span data-ttu-id="c3a4a-601">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-601">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-602">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-602">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-603">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-603">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-604">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-604">Compose or Read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="c3a4a-605">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-605">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="c3a4a-606">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-606">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="c3a4a-607">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-607">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a4a-608">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-608">Read mode</span></span>

<span data-ttu-id="c3a4a-609">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-609">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a4a-610">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-610">Compose mode</span></span>

<span data-ttu-id="c3a4a-611">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-611">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="c3a4a-612">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-612">Type</span></span>

*   <span data-ttu-id="c3a4a-613">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-613">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-614">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-614">Requirements</span></span>

|<span data-ttu-id="c3a4a-615">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-615">Requirement</span></span>|<span data-ttu-id="c3a4a-616">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-616">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-617">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-617">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-618">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-618">1.0</span></span>|
|[<span data-ttu-id="c3a4a-619">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-619">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-620">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-620">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-621">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-621">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-622">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-622">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="c3a4a-623">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-623">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="c3a4a-p128">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="c3a4a-p129">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p129">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a4a-628">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-628">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a4a-629">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-629">Type</span></span>

*   [<span data-ttu-id="c3a4a-630">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c3a4a-630">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="c3a4a-631">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-631">Requirements</span></span>

|<span data-ttu-id="c3a4a-632">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-632">Requirement</span></span>|<span data-ttu-id="c3a4a-633">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-633">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-634">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-634">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-635">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-635">1.0</span></span>|
|[<span data-ttu-id="c3a4a-636">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-636">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-637">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-637">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-638">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-638">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-639">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-639">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-640">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-640">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="c3a4a-641">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-641">(nullable) seriesId :String</span></span>

<span data-ttu-id="c3a4a-642">あるインスタンスが属するシリーズの ID を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-642">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="c3a4a-643">OWA と Outlook では、`seriesId` はこのアイテムが属する親 (シリーズ) アイテムの Exchange Web Services (EWS) ID を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-643">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="c3a4a-644">ただし、iOS と Android の場合、`seriesId` は親アイテムの REST ID を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-644">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a4a-645">`seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-645">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="c3a4a-646">`seriesId` プロパティは、Outlook REST API で使用される Outlook ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-646">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="c3a4a-647">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-647">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="c3a4a-648">詳細については、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-648">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="c3a4a-649">`seriesId` プロパティは、単発の予定、シリーズ アイテム、会議出席依頼など、親アイテムを持たないアイテムに対して `null` を返し、会議出席依頼ではないその他のアイテムに対して `undefined` を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-649">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a4a-650">Type</span><span class="sxs-lookup"><span data-stu-id="c3a4a-650">Type</span></span>

* <span data-ttu-id="c3a4a-651">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-651">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-652">Requirements</span><span class="sxs-lookup"><span data-stu-id="c3a4a-652">Requirements</span></span>

|<span data-ttu-id="c3a4a-653">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-653">Requirement</span></span>|<span data-ttu-id="c3a4a-654">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-654">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-655">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-655">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-656">1.7</span><span class="sxs-lookup"><span data-stu-id="c3a4a-656">1.7</span></span>|
|[<span data-ttu-id="c3a4a-657">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-657">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-658">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-658">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-659">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-659">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-660">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-660">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-661">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-661">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="c3a4a-662">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-662">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="c3a4a-663">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-663">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="c3a4a-p132">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a4a-666">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-666">Read mode</span></span>

<span data-ttu-id="c3a4a-667">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-667">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a4a-668">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-668">Compose mode</span></span>

<span data-ttu-id="c3a4a-669">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-669">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="c3a4a-670">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-670">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="c3a4a-671">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-671">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="c3a4a-672">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-672">Type</span></span>

*   <span data-ttu-id="c3a4a-673">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-673">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-674">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-674">Requirements</span></span>

|<span data-ttu-id="c3a4a-675">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-675">Requirement</span></span>|<span data-ttu-id="c3a4a-676">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-676">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-677">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-677">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-678">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-678">1.0</span></span>|
|[<span data-ttu-id="c3a4a-679">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-679">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-680">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-680">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-681">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-681">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-682">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-682">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="c3a4a-683">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-683">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="c3a4a-684">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-684">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="c3a4a-685">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-685">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a4a-686">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-686">Read mode</span></span>

<span data-ttu-id="c3a4a-p133">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="c3a4a-689">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-689">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a4a-690">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-690">Compose mode</span></span>
<span data-ttu-id="c3a4a-691">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-691">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="c3a4a-692">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-692">Type</span></span>

*   <span data-ttu-id="c3a4a-693">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-693">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-694">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-694">Requirements</span></span>

|<span data-ttu-id="c3a4a-695">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-695">Requirement</span></span>|<span data-ttu-id="c3a4a-696">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-696">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-697">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-697">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-698">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-698">1.0</span></span>|
|[<span data-ttu-id="c3a4a-699">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-699">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-700">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-700">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-701">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-701">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-702">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-702">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="c3a4a-703">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-703">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="c3a4a-704">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-704">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="c3a4a-705">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-705">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a4a-706">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-706">Read mode</span></span>

<span data-ttu-id="c3a4a-p135">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a4a-709">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-709">Compose mode</span></span>

<span data-ttu-id="c3a4a-710">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-710">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c3a4a-711">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-711">Type</span></span>

*   <span data-ttu-id="c3a4a-712">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-712">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-713">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-713">Requirements</span></span>

|<span data-ttu-id="c3a4a-714">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-714">Requirement</span></span>|<span data-ttu-id="c3a4a-715">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-715">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-716">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-716">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-717">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-717">1.0</span></span>|
|[<span data-ttu-id="c3a4a-718">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-718">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-719">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-719">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-720">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-720">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-721">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-721">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="c3a4a-722">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a4a-722">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="c3a4a-723">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c3a4a-723">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c3a4a-724">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-724">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="c3a4a-725">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-725">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="c3a4a-726">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-726">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a4a-727">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a4a-727">Parameters</span></span>
|<span data-ttu-id="c3a4a-728">名前</span><span class="sxs-lookup"><span data-stu-id="c3a4a-728">Name</span></span>|<span data-ttu-id="c3a4a-729">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-729">Type</span></span>|<span data-ttu-id="c3a4a-730">属性</span><span class="sxs-lookup"><span data-stu-id="c3a4a-730">Attributes</span></span>|<span data-ttu-id="c3a4a-731">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-731">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="c3a4a-732">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-732">String</span></span>||<span data-ttu-id="c3a4a-p136">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="c3a4a-735">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-735">String</span></span>||<span data-ttu-id="c3a4a-p137">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="c3a4a-738">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-738">Object</span></span>|<span data-ttu-id="c3a4a-739">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-739">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-740">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-740">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a4a-741">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-741">Object</span></span>|<span data-ttu-id="c3a4a-742">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-742">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-743">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-743">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="c3a4a-744">Boolean</span><span class="sxs-lookup"><span data-stu-id="c3a4a-744">Boolean</span></span>|<span data-ttu-id="c3a4a-745">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-745">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-746">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-746">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="c3a4a-747">function</span><span class="sxs-lookup"><span data-stu-id="c3a4a-747">function</span></span>|<span data-ttu-id="c3a4a-748">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-748">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-749">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-749">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c3a4a-750">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-750">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c3a4a-751">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-751">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c3a4a-752">エラー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-752">Errors</span></span>

|<span data-ttu-id="c3a4a-753">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-753">Error code</span></span>|<span data-ttu-id="c3a4a-754">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-754">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="c3a4a-755">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-755">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="c3a4a-756">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-756">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="c3a4a-757">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-757">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a4a-758">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-758">Requirements</span></span>

|<span data-ttu-id="c3a4a-759">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-759">Requirement</span></span>|<span data-ttu-id="c3a4a-760">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-761">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-761">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-762">1.1</span><span class="sxs-lookup"><span data-stu-id="c3a4a-762">1.1</span></span>|
|[<span data-ttu-id="c3a4a-763">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-763">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-764">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-764">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3a4a-765">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-765">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-766">作成</span><span class="sxs-lookup"><span data-stu-id="c3a4a-766">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c3a4a-767">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-767">Examples</span></span>

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

<span data-ttu-id="c3a4a-768">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-768">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="c3a4a-769">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c3a4a-769">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c3a4a-770">ファイルを添付ファイルとして base64 エンコーディングからメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-770">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="c3a4a-771">`addFileAttachmentFromBase64Async` メソッドは、base64 エンコーディングからファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-771">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="c3a4a-772">このメソッドによって、AsyncResult.value オブジェクトの添付ファイル識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-772">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="c3a4a-773">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-773">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a4a-774">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a4a-774">Parameters</span></span>
|<span data-ttu-id="c3a4a-775">名前</span><span class="sxs-lookup"><span data-stu-id="c3a4a-775">Name</span></span>|<span data-ttu-id="c3a4a-776">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-776">Type</span></span>|<span data-ttu-id="c3a4a-777">属性</span><span class="sxs-lookup"><span data-stu-id="c3a4a-777">Attributes</span></span>|<span data-ttu-id="c3a4a-778">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-778">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="c3a4a-779">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-779">String</span></span>||<span data-ttu-id="c3a4a-780">電子メールまたはイベントに追加する画像またはファイルの base64 でエンコードされたコンテンツ。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-780">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="c3a4a-781">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-781">String</span></span>||<span data-ttu-id="c3a4a-p139">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p139">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="c3a4a-784">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-784">Object</span></span>|<span data-ttu-id="c3a4a-785">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-785">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-786">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-786">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a4a-787">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-787">Object</span></span>|<span data-ttu-id="c3a4a-788">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-788">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-789">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-789">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="c3a4a-790">Boolean</span><span class="sxs-lookup"><span data-stu-id="c3a4a-790">Boolean</span></span>|<span data-ttu-id="c3a4a-791">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-791">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-792">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-792">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="c3a4a-793">function</span><span class="sxs-lookup"><span data-stu-id="c3a4a-793">function</span></span>|<span data-ttu-id="c3a4a-794">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-794">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-795">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-795">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c3a4a-796">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-796">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c3a4a-797">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-797">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c3a4a-798">エラー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-798">Errors</span></span>

|<span data-ttu-id="c3a4a-799">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-799">Error code</span></span>|<span data-ttu-id="c3a4a-800">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-800">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="c3a4a-801">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-801">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="c3a4a-802">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-802">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="c3a4a-803">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-803">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a4a-804">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-804">Requirements</span></span>

|<span data-ttu-id="c3a4a-805">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-805">Requirement</span></span>|<span data-ttu-id="c3a4a-806">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-806">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-807">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-807">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-808">プレビュー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-808">Preview</span></span>|
|[<span data-ttu-id="c3a4a-809">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-809">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-810">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-810">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3a4a-811">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-811">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-812">新規作成</span><span class="sxs-lookup"><span data-stu-id="c3a4a-812">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c3a4a-813">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-813">Examples</span></span>

```javascript
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
        // Do something here.
      });
  });
```

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="c3a4a-814">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c3a4a-814">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="c3a4a-815">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-815">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="c3a4a-816">現在、サポートされているイベントの種類は `Office.EventType.AttachmentsChanged`、`Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged`、`Office.EventType.RecurrenceChanged` です。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-816">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a4a-817">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a4a-817">Parameters</span></span>

| <span data-ttu-id="c3a4a-818">名前</span><span class="sxs-lookup"><span data-stu-id="c3a4a-818">Name</span></span> | <span data-ttu-id="c3a4a-819">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-819">Type</span></span> | <span data-ttu-id="c3a4a-820">属性</span><span class="sxs-lookup"><span data-stu-id="c3a4a-820">Attributes</span></span> | <span data-ttu-id="c3a4a-821">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-821">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c3a4a-822">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c3a4a-822">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c3a4a-823">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-823">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="c3a4a-824">Function</span><span class="sxs-lookup"><span data-stu-id="c3a4a-824">Function</span></span> || <span data-ttu-id="c3a4a-p140">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p140">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="c3a4a-828">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-828">Object</span></span> | <span data-ttu-id="c3a4a-829">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-829">&lt;optional&gt;</span></span> | <span data-ttu-id="c3a4a-830">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-830">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c3a4a-831">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-831">Object</span></span> | <span data-ttu-id="c3a4a-832">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-832">&lt;optional&gt;</span></span> | <span data-ttu-id="c3a4a-833">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-833">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c3a4a-834">function</span><span class="sxs-lookup"><span data-stu-id="c3a4a-834">function</span></span>| <span data-ttu-id="c3a4a-835">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-835">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-836">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-836">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a4a-837">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-837">Requirements</span></span>

|<span data-ttu-id="c3a4a-838">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-838">Requirement</span></span>| <span data-ttu-id="c3a4a-839">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-839">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-840">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-840">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3a4a-841">1.7</span><span class="sxs-lookup"><span data-stu-id="c3a4a-841">1.7</span></span> |
|[<span data-ttu-id="c3a4a-842">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-842">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3a4a-843">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-843">ReadItem</span></span> |
|[<span data-ttu-id="c3a4a-844">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-844">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3a4a-845">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-845">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="c3a4a-846">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-846">Example</span></span>

```javascript
function myHandlerFunction(eventarg) {
  if (eventarg.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added) {
    var attachment = eventarg.attachmentDetails;
    console.log("Event Fired and Attachment Added!");
    getAttachmentContentAsync(attachment.id, options, callback);
  }
}

Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, myHandlerFunction, myCallback);
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="c3a4a-847">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c3a4a-847">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c3a4a-848">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-848">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="c3a4a-p141">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="c3a4a-852">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-852">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="c3a4a-853">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-853">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a4a-854">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a4a-854">Parameters</span></span>

|<span data-ttu-id="c3a4a-855">名前</span><span class="sxs-lookup"><span data-stu-id="c3a4a-855">Name</span></span>|<span data-ttu-id="c3a4a-856">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-856">Type</span></span>|<span data-ttu-id="c3a4a-857">属性</span><span class="sxs-lookup"><span data-stu-id="c3a4a-857">Attributes</span></span>|<span data-ttu-id="c3a4a-858">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-858">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="c3a4a-859">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-859">String</span></span>||<span data-ttu-id="c3a4a-p142">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="c3a4a-862">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-862">String</span></span>||<span data-ttu-id="c3a4a-863">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-863">The subject of the item to be attached.</span></span> <span data-ttu-id="c3a4a-864">最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-864">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="c3a4a-865">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-865">Object</span></span>|<span data-ttu-id="c3a4a-866">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-866">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-867">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-867">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a4a-868">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-868">Object</span></span>|<span data-ttu-id="c3a4a-869">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-869">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-870">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-870">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c3a4a-871">function</span><span class="sxs-lookup"><span data-stu-id="c3a4a-871">function</span></span>|<span data-ttu-id="c3a4a-872">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-872">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-873">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-873">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c3a4a-874">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-874">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c3a4a-875">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-875">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c3a4a-876">エラー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-876">Errors</span></span>

|<span data-ttu-id="c3a4a-877">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-877">Error code</span></span>|<span data-ttu-id="c3a4a-878">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-878">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="c3a4a-879">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-879">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a4a-880">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-880">Requirements</span></span>

|<span data-ttu-id="c3a4a-881">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-881">Requirement</span></span>|<span data-ttu-id="c3a4a-882">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-882">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-883">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-883">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-884">1.1</span><span class="sxs-lookup"><span data-stu-id="c3a4a-884">1.1</span></span>|
|[<span data-ttu-id="c3a4a-885">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-885">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-886">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-886">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3a4a-887">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-887">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-888">作成</span><span class="sxs-lookup"><span data-stu-id="c3a4a-888">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-889">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-889">Example</span></span>

<span data-ttu-id="c3a4a-890">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-890">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="c3a4a-891">close()</span><span class="sxs-lookup"><span data-stu-id="c3a4a-891">close()</span></span>

<span data-ttu-id="c3a4a-892">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-892">Closes the current item that is being composed.</span></span>

<span data-ttu-id="c3a4a-p144">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a4a-895">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-895">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="c3a4a-896">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-896">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-897">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-897">Requirements</span></span>

|<span data-ttu-id="c3a4a-898">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-898">Requirement</span></span>|<span data-ttu-id="c3a4a-899">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-899">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-900">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-900">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-901">1.3</span><span class="sxs-lookup"><span data-stu-id="c3a4a-901">1.3</span></span>|
|[<span data-ttu-id="c3a4a-902">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-902">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-903">制限あり</span><span class="sxs-lookup"><span data-stu-id="c3a4a-903">Restricted</span></span>|
|[<span data-ttu-id="c3a4a-904">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-904">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-905">作成</span><span class="sxs-lookup"><span data-stu-id="c3a4a-905">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="c3a4a-906">displayReplyAllForm (formdata, [callback])</span><span class="sxs-lookup"><span data-stu-id="c3a4a-906">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="c3a4a-907">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-907">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a4a-908">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-908">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c3a4a-909">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-909">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c3a4a-910">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-910">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="c3a4a-p145">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a4a-914">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a4a-914">Parameters</span></span>

|<span data-ttu-id="c3a4a-915">名前</span><span class="sxs-lookup"><span data-stu-id="c3a4a-915">Name</span></span>|<span data-ttu-id="c3a4a-916">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-916">Type</span></span>|<span data-ttu-id="c3a4a-917">属性</span><span class="sxs-lookup"><span data-stu-id="c3a4a-917">Attributes</span></span>|<span data-ttu-id="c3a4a-918">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-918">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="c3a4a-919">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-919">String &#124; Object</span></span>||<span data-ttu-id="c3a4a-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c3a4a-922">**または**</span><span class="sxs-lookup"><span data-stu-id="c3a4a-922">**OR**</span></span><br/><span data-ttu-id="c3a4a-p147">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="c3a4a-925">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-925">String</span></span>|<span data-ttu-id="c3a4a-926">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-926">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-p148">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="c3a4a-929">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-929">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="c3a4a-930">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-930">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-931">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-931">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="c3a4a-932">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-932">String</span></span>||<span data-ttu-id="c3a4a-p149">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="c3a4a-935">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-935">String</span></span>||<span data-ttu-id="c3a4a-936">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-936">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="c3a4a-937">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-937">String</span></span>||<span data-ttu-id="c3a4a-p150">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="c3a4a-940">Boolean</span><span class="sxs-lookup"><span data-stu-id="c3a4a-940">Boolean</span></span>||<span data-ttu-id="c3a4a-p151">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="c3a4a-943">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-943">String</span></span>||<span data-ttu-id="c3a4a-p152">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="c3a4a-947">function</span><span class="sxs-lookup"><span data-stu-id="c3a4a-947">function</span></span>|<span data-ttu-id="c3a4a-948">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-948">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-949">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-949">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a4a-950">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-950">Requirements</span></span>

|<span data-ttu-id="c3a4a-951">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-951">Requirement</span></span>|<span data-ttu-id="c3a4a-952">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-952">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-953">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-953">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-954">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-954">1.0</span></span>|
|[<span data-ttu-id="c3a4a-955">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-955">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-956">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-956">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-957">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-957">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-958">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-958">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c3a4a-959">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-959">Examples</span></span>

<span data-ttu-id="c3a4a-960">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-960">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="c3a4a-961">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-961">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="c3a4a-962">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-962">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c3a4a-963">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-963">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c3a4a-964">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-964">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c3a4a-965">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-965">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="c3a4a-966">displayReplyForm (formdata, [callback])</span><span class="sxs-lookup"><span data-stu-id="c3a4a-966">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="c3a4a-967">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-967">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a4a-968">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-968">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c3a4a-969">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-969">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c3a4a-970">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-970">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="c3a4a-p153">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p153">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a4a-974">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a4a-974">Parameters</span></span>

|<span data-ttu-id="c3a4a-975">名前</span><span class="sxs-lookup"><span data-stu-id="c3a4a-975">Name</span></span>|<span data-ttu-id="c3a4a-976">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-976">Type</span></span>|<span data-ttu-id="c3a4a-977">属性</span><span class="sxs-lookup"><span data-stu-id="c3a4a-977">Attributes</span></span>|<span data-ttu-id="c3a4a-978">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-978">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="c3a4a-979">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-979">String &#124; Object</span></span>||<span data-ttu-id="c3a4a-p154">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c3a4a-982">**または**</span><span class="sxs-lookup"><span data-stu-id="c3a4a-982">**OR**</span></span><br/><span data-ttu-id="c3a4a-p155">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="c3a4a-985">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-985">String</span></span>|<span data-ttu-id="c3a4a-986">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-986">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-p156">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="c3a4a-989">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-989">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="c3a4a-990">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-990">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-991">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-991">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="c3a4a-992">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-992">String</span></span>||<span data-ttu-id="c3a4a-p157">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="c3a4a-995">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-995">String</span></span>||<span data-ttu-id="c3a4a-996">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-996">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="c3a4a-997">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-997">String</span></span>||<span data-ttu-id="c3a4a-p158">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="c3a4a-1000">Boolean</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1000">Boolean</span></span>||<span data-ttu-id="c3a4a-p159">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="c3a4a-1003">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1003">String</span></span>||<span data-ttu-id="c3a4a-p160">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="c3a4a-1007">function</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1007">function</span></span>|<span data-ttu-id="c3a4a-1008">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1008">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1009">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1009">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a4a-1010">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1010">Requirements</span></span>

|<span data-ttu-id="c3a4a-1011">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1011">Requirement</span></span>|<span data-ttu-id="c3a4a-1012">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1012">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-1013">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1013">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-1014">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1014">1.0</span></span>|
|[<span data-ttu-id="c3a4a-1015">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1015">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-1016">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1016">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-1017">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1017">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-1018">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1018">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c3a4a-1019">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1019">Examples</span></span>

<span data-ttu-id="c3a4a-1020">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1020">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="c3a4a-1021">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1021">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="c3a4a-1022">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1022">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c3a4a-1023">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1023">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c3a4a-1024">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1024">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c3a4a-1025">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1025">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="c3a4a-1026">getattachmentcontentasync (attachmentId, [options], [callback]) > [attachmentcontent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1026">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="c3a4a-1027">メッセージまたは予定から指定の添付ファイルを取得し、それを `AttachmentContent` オブジェクトとして返されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1027">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="c3a4a-1028">`getAttachmentContentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1028">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="c3a4a-1029">ベスト プラクティスとして、識別子を使用し、`getAttachmentsAsync` または `item.attachments` 呼び出しで attachmentIds を取得した同じセッションで添付ファイルを取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1029">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="c3a4a-1030">Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1030">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="c3a4a-1031">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始し、フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1031">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a4a-1032">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1032">Parameters</span></span>

|<span data-ttu-id="c3a4a-1033">名前</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1033">Name</span></span>|<span data-ttu-id="c3a4a-1034">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1034">Type</span></span>|<span data-ttu-id="c3a4a-1035">属性</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1035">Attributes</span></span>|<span data-ttu-id="c3a4a-1036">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1036">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="c3a4a-1037">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1037">String</span></span>||<span data-ttu-id="c3a4a-1038">取得する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1038">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="c3a4a-1039">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1039">Object</span></span>|<span data-ttu-id="c3a4a-1040">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1040">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1041">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1041">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a4a-1042">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1042">Object</span></span>|<span data-ttu-id="c3a4a-1043">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1043">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1044">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1044">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c3a4a-1045">function</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1045">function</span></span>|<span data-ttu-id="c3a4a-1046">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1046">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1047">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1047">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a4a-1048">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1048">Requirements</span></span>

|<span data-ttu-id="c3a4a-1049">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1049">Requirement</span></span>|<span data-ttu-id="c3a4a-1050">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-1051">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1051">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-1052">プレビュー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1052">Preview</span></span>|
|[<span data-ttu-id="c3a4a-1053">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1053">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-1054">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1054">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-1055">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1055">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-1056">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1056">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3a4a-1057">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1057">Returns:</span></span>

<span data-ttu-id="c3a4a-1058">型: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1058">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="c3a4a-1059">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1059">Example</span></span>

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
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  if (result.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
    // Handle file attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.Eml) {
    // Handle email item attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
    // Handle .icalender attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.Url) {
    // Handle cloud attachment.
  } else {
    // Handle attachment formats that are not supported.
  }
}
```

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="c3a4a-1060">getAttachmentsAsync ([オプション], [callback])] > <[attachmentdetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c3a4a-1060">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="c3a4a-1061">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1061">Gets the item's attachments as an array.</span></span> <span data-ttu-id="c3a4a-1062">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1062">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a4a-1063">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1063">Parameters</span></span>

|<span data-ttu-id="c3a4a-1064">名前</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1064">Name</span></span>|<span data-ttu-id="c3a4a-1065">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1065">Type</span></span>|<span data-ttu-id="c3a4a-1066">属性</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1066">Attributes</span></span>|<span data-ttu-id="c3a4a-1067">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1067">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c3a4a-1068">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1068">Object</span></span>|<span data-ttu-id="c3a4a-1069">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1070">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1070">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a4a-1071">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1071">Object</span></span>|<span data-ttu-id="c3a4a-1072">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1072">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1073">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1073">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c3a4a-1074">function</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1074">function</span></span>|<span data-ttu-id="c3a4a-1075">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1075">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1076">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1076">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a4a-1077">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1077">Requirements</span></span>

|<span data-ttu-id="c3a4a-1078">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1078">Requirement</span></span>|<span data-ttu-id="c3a4a-1079">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1079">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-1080">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1080">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-1081">プレビュー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1081">Preview</span></span>|
|[<span data-ttu-id="c3a4a-1082">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1082">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-1083">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1083">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-1084">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1084">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-1085">作成</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1085">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3a4a-1086">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1086">Returns:</span></span>

<span data-ttu-id="c3a4a-1087">型: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c3a4a-1087">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="c3a4a-1088">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1088">Example</span></span>

<span data-ttu-id="c3a4a-1089">次の例では、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1089">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      var attachment = result.value [i];
      outputString += "<BR>" + i + ". Name: ";
      outputString += attachment.name;
      outputString += "<BR>ID: " + attachment.id;
      outputString += "<BR>contentType: " + attachment.contentType;
      outputString += "<BR>size: " + attachment.size;
      outputString += "<BR>attachmentType: " + attachment.attachmentType;
      outputString += "<BR>isInline: " + attachment.isInline;
    }
  }
}
```

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="c3a4a-1090">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1090">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="c3a4a-1091">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1091">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a4a-1092">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1092">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-1093">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1093">Requirements</span></span>

|<span data-ttu-id="c3a4a-1094">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1094">Requirement</span></span>|<span data-ttu-id="c3a4a-1095">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1095">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-1096">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1096">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-1097">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1097">1.0</span></span>|
|[<span data-ttu-id="c3a4a-1098">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1098">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-1099">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1099">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-1100">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1100">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-1101">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1101">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3a4a-1102">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1102">Returns:</span></span>

<span data-ttu-id="c3a4a-1103">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1103">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="c3a4a-1104">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1104">Example</span></span>

<span data-ttu-id="c3a4a-1105">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1105">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="c3a4a-1106">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1106">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c3a4a-1107">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1107">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a4a-1108">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1108">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a4a-1109">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1109">Parameters</span></span>

|<span data-ttu-id="c3a4a-1110">名前</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1110">Name</span></span>|<span data-ttu-id="c3a4a-1111">種類</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1111">Type</span></span>|<span data-ttu-id="c3a4a-1112">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1112">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="c3a4a-1113">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1113">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="c3a4a-1114">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1114">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a4a-1115">Requirements</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1115">Requirements</span></span>

|<span data-ttu-id="c3a4a-1116">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1116">Requirement</span></span>|<span data-ttu-id="c3a4a-1117">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1117">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-1118">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1118">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-1119">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1119">1.0</span></span>|
|[<span data-ttu-id="c3a4a-1120">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1120">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-1121">制限あり</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1121">Restricted</span></span>|
|[<span data-ttu-id="c3a4a-1122">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1122">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-1123">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1123">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3a4a-1124">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1124">Returns:</span></span>

<span data-ttu-id="c3a4a-1125">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1125">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="c3a4a-1126">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1126">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="c3a4a-1127">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1127">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="c3a4a-1128">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1128">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="c3a4a-1129">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1129">Value of `entityType`</span></span>|<span data-ttu-id="c3a4a-1130">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1130">Type of objects in returned array</span></span>|<span data-ttu-id="c3a4a-1131">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1131">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="c3a4a-1132">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1132">String</span></span>|<span data-ttu-id="c3a4a-1133">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1133">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="c3a4a-1134">連絡先</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1134">Contact</span></span>|<span data-ttu-id="c3a4a-1135">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1135">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="c3a4a-1136">文字列</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1136">String</span></span>|<span data-ttu-id="c3a4a-1137">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1137">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="c3a4a-1138">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1138">MeetingSuggestion</span></span>|<span data-ttu-id="c3a4a-1139">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1139">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="c3a4a-1140">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1140">PhoneNumber</span></span>|<span data-ttu-id="c3a4a-1141">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1141">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="c3a4a-1142">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1142">TaskSuggestion</span></span>|<span data-ttu-id="c3a4a-1143">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1143">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="c3a4a-1144">文字列</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1144">String</span></span>|<span data-ttu-id="c3a4a-1145">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1145">**Restricted**</span></span>|

<span data-ttu-id="c3a4a-1146">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c3a4a-1146">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="c3a4a-1147">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1147">Example</span></span>

<span data-ttu-id="c3a4a-1148">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1148">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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
};
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="c3a4a-1149">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1149">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c3a4a-1150">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1150">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a4a-1151">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1151">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c3a4a-1152">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1152">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a4a-1153">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1153">Parameters</span></span>

|<span data-ttu-id="c3a4a-1154">名前</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1154">Name</span></span>|<span data-ttu-id="c3a4a-1155">種類</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1155">Type</span></span>|<span data-ttu-id="c3a4a-1156">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1156">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="c3a4a-1157">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1157">String</span></span>|<span data-ttu-id="c3a4a-1158">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1158">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a4a-1159">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1159">Requirements</span></span>

|<span data-ttu-id="c3a4a-1160">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1160">Requirement</span></span>|<span data-ttu-id="c3a4a-1161">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1161">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-1162">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1162">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-1163">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1163">1.0</span></span>|
|[<span data-ttu-id="c3a4a-1164">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1164">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-1165">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1165">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-1166">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1166">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-1167">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1167">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3a4a-1168">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1168">Returns:</span></span>

<span data-ttu-id="c3a4a-p164">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p164">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="c3a4a-1171">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c3a4a-1171">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="c3a4a-1172">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1172">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="c3a4a-1173">アドインが[操作可能メッセージによってアクティブ化](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1173">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="c3a4a-1174">このメソッドは、Outlook 2016 for Windows 以降 (16.0.8413.1000 以降のクイック実行バージョン) および Outlook on the web for Office 365 でのみサポートされます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1174">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a4a-1175">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1175">Parameters</span></span>
|<span data-ttu-id="c3a4a-1176">名前</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1176">Name</span></span>|<span data-ttu-id="c3a4a-1177">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1177">Type</span></span>|<span data-ttu-id="c3a4a-1178">属性</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1178">Attributes</span></span>|<span data-ttu-id="c3a4a-1179">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1179">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c3a4a-1180">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1180">Object</span></span>|<span data-ttu-id="c3a4a-1181">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1181">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1182">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1182">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a4a-1183">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1183">Object</span></span>|<span data-ttu-id="c3a4a-1184">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1184">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1185">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1185">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c3a4a-1186">function</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1186">function</span></span>|<span data-ttu-id="c3a4a-1187">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1187">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1188">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1188">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c3a4a-1189">成功すると、初期化データが文字列として `asyncResult.value` プロパティで指定されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1189">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="c3a4a-1190">初期化コンテキストがない場合、`asyncResult` オブジェクトには、`code` プロパティが `9020`、`name` プロパティが `GenericResponseError` に設定された `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1190">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a4a-1191">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1191">Requirements</span></span>

|<span data-ttu-id="c3a4a-1192">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1192">Requirement</span></span>|<span data-ttu-id="c3a4a-1193">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-1194">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-1195">プレビュー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1195">Preview</span></span>|
|[<span data-ttu-id="c3a4a-1196">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-1197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1197">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-1198">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-1199">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1199">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-1200">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1200">Example</span></span>

```javascript
// Get the initialization context (if present).
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object.
        var context = JSON.parse(asyncResult.value);
        // Do something with context.
      } else {
        // Empty context, treat as no context.
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is no context.
        // Treat as no context.
      } else {
        // Handle the error.
      }
    }
  }
);
```

#### <a name="getregexmatches--object"></a><span data-ttu-id="c3a4a-1201">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1201">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="c3a4a-1202">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1202">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a4a-1203">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1203">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c3a4a-p165">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p165">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c3a4a-1207">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1207">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c3a4a-1208">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1208">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c3a4a-p166">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-1212">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1212">Requirements</span></span>

|<span data-ttu-id="c3a4a-1213">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1213">Requirement</span></span>|<span data-ttu-id="c3a4a-1214">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1214">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-1215">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1215">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-1216">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1216">1.0</span></span>|
|[<span data-ttu-id="c3a4a-1217">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1217">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-1218">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1218">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-1219">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1219">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-1220">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1220">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3a4a-1221">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1221">Returns:</span></span>

<span data-ttu-id="c3a4a-p167">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="c3a4a-1224">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1224">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c3a4a-1225">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1225">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c3a4a-1226">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1226">Example</span></span>

<span data-ttu-id="c3a4a-1227">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1227">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="c3a4a-1228">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1228">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="c3a4a-1229">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1229">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a4a-1230">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1230">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c3a4a-1231">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1231">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="c3a4a-p168">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p168">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a4a-1234">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1234">Parameters</span></span>

|<span data-ttu-id="c3a4a-1235">名前</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1235">Name</span></span>|<span data-ttu-id="c3a4a-1236">種類</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1236">Type</span></span>|<span data-ttu-id="c3a4a-1237">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1237">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="c3a4a-1238">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1238">String</span></span>|<span data-ttu-id="c3a4a-1239">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1239">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a4a-1240">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1240">Requirements</span></span>

|<span data-ttu-id="c3a4a-1241">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1241">Requirement</span></span>|<span data-ttu-id="c3a4a-1242">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-1243">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1243">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-1244">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1244">1.0</span></span>|
|[<span data-ttu-id="c3a4a-1245">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1245">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1246">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-1247">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1247">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-1248">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1248">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3a4a-1249">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1249">Returns:</span></span>

<span data-ttu-id="c3a4a-1250">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1250">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="c3a4a-1251">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1251">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c3a4a-1252">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="c3a4a-1252">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c3a4a-1253">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1253">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="c3a4a-1254">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1254">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="c3a4a-1255">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1255">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="c3a4a-p169">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p169">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a4a-1258">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1258">Parameters</span></span>

|<span data-ttu-id="c3a4a-1259">名前</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1259">Name</span></span>|<span data-ttu-id="c3a4a-1260">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1260">Type</span></span>|<span data-ttu-id="c3a4a-1261">属性</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1261">Attributes</span></span>|<span data-ttu-id="c3a4a-1262">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1262">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="c3a4a-1263">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1263">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="c3a4a-p170">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p170">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="c3a4a-1267">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1267">Object</span></span>|<span data-ttu-id="c3a4a-1268">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1268">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1269">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1269">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a4a-1270">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1270">Object</span></span>|<span data-ttu-id="c3a4a-1271">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1271">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1272">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1272">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c3a4a-1273">function</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1273">function</span></span>||<span data-ttu-id="c3a4a-1274">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1274">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c3a4a-1275">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1275">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="c3a4a-1276">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1276">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a4a-1277">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1277">Requirements</span></span>

|<span data-ttu-id="c3a4a-1278">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1278">Requirement</span></span>|<span data-ttu-id="c3a4a-1279">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1279">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-1280">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-1281">1.2</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1281">1.2</span></span>|
|[<span data-ttu-id="c3a4a-1282">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-1283">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1283">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3a4a-1284">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-1285">作成</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1285">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3a4a-1286">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1286">Returns:</span></span>

<span data-ttu-id="c3a4a-1287">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1287">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="c3a4a-1288">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1288">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c3a4a-1289">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1289">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c3a4a-1290">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1290">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="c3a4a-1291">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1291">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="c3a4a-p172">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p172">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="c3a4a-1294">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1294">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-1295">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1295">Requirements</span></span>

|<span data-ttu-id="c3a4a-1296">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1296">Requirement</span></span>|<span data-ttu-id="c3a4a-1297">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1297">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-1298">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-1299">1.6</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1299">1.6</span></span>|
|[<span data-ttu-id="c3a4a-1300">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1300">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-1301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1301">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-1302">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1302">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-1303">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1303">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3a4a-1304">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1304">Returns:</span></span>

<span data-ttu-id="c3a4a-1305">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1305">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="c3a4a-1306">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1306">Example</span></span>

<span data-ttu-id="c3a4a-1307">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1307">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="c3a4a-1308">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1308">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="c3a4a-p173">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p173">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="c3a4a-1311">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1311">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c3a4a-p174">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p174">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c3a4a-1315">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1315">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c3a4a-1316">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1316">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c3a4a-p175">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p175">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a4a-1320">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1320">Requirements</span></span>

|<span data-ttu-id="c3a4a-1321">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1321">Requirement</span></span>|<span data-ttu-id="c3a4a-1322">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1322">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-1323">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1323">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-1324">1.6</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1324">1.6</span></span>|
|[<span data-ttu-id="c3a4a-1325">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-1326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1326">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-1327">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-1328">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1328">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3a4a-1329">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1329">Returns:</span></span>

<span data-ttu-id="c3a4a-p176">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p176">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="c3a4a-1332">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1332">Example</span></span>

<span data-ttu-id="c3a4a-1333">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1333">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="c3a4a-1334">getSharedPropertiesAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1334">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="c3a4a-1335">共有フォルダー、カレンダー、メールボックスで選択した予定またはメッセージのプロパティを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1335">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a4a-1336">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1336">Parameters</span></span>

|<span data-ttu-id="c3a4a-1337">名前</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1337">Name</span></span>|<span data-ttu-id="c3a4a-1338">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1338">Type</span></span>|<span data-ttu-id="c3a4a-1339">属性</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1339">Attributes</span></span>|<span data-ttu-id="c3a4a-1340">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1340">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c3a4a-1341">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1341">Object</span></span>|<span data-ttu-id="c3a4a-1342">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1342">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1343">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1343">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a4a-1344">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1344">Object</span></span>|<span data-ttu-id="c3a4a-1345">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1345">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1346">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1346">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c3a4a-1347">function</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1347">function</span></span>||<span data-ttu-id="c3a4a-1348">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1348">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c3a4a-1349">共有プロパティは `asyncResult.value` プロパティの [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1349">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="c3a4a-1350">このオブジェクトは、アイテムの共有プロパティの取得に使用できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1350">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a4a-1351">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1351">Requirements</span></span>

|<span data-ttu-id="c3a4a-1352">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1352">Requirement</span></span>|<span data-ttu-id="c3a4a-1353">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1353">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-1354">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1354">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-1355">プレビュー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1355">Preview</span></span>|
|[<span data-ttu-id="c3a4a-1356">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1356">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-1357">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1357">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-1358">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1358">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-1359">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1359">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-1360">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1360">Example</span></span>

```javascript
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="c3a4a-1361">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1361">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="c3a4a-1362">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1362">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="c3a4a-p178">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p178">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a4a-1366">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1366">Parameters</span></span>

|<span data-ttu-id="c3a4a-1367">名前</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1367">Name</span></span>|<span data-ttu-id="c3a4a-1368">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1368">Type</span></span>|<span data-ttu-id="c3a4a-1369">属性</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1369">Attributes</span></span>|<span data-ttu-id="c3a4a-1370">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1370">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="c3a4a-1371">function</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1371">function</span></span>||<span data-ttu-id="c3a4a-1372">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1372">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c3a4a-1373">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1373">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="c3a4a-1374">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1374">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="c3a4a-1375">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1375">Object</span></span>|<span data-ttu-id="c3a4a-1376">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1376">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1377">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1377">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="c3a4a-1378">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1378">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a4a-1379">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1379">Requirements</span></span>

|<span data-ttu-id="c3a4a-1380">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1380">Requirement</span></span>|<span data-ttu-id="c3a4a-1381">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1381">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-1382">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1382">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-1383">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1383">1.0</span></span>|
|[<span data-ttu-id="c3a4a-1384">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1384">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-1385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1385">ReadItem</span></span>|
|[<span data-ttu-id="c3a4a-1386">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1386">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-1387">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1387">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-1388">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1388">Example</span></span>

<span data-ttu-id="c3a4a-p181">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p181">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="c3a4a-1392">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1392">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="c3a4a-1393">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1393">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="c3a4a-1394">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1394">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="c3a4a-1395">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1395">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="c3a4a-1396">Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1396">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="c3a4a-1397">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始し、フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1397">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a4a-1398">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1398">Parameters</span></span>

|<span data-ttu-id="c3a4a-1399">名前</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1399">Name</span></span>|<span data-ttu-id="c3a4a-1400">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1400">Type</span></span>|<span data-ttu-id="c3a4a-1401">属性</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1401">Attributes</span></span>|<span data-ttu-id="c3a4a-1402">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1402">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="c3a4a-1403">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1403">String</span></span>||<span data-ttu-id="c3a4a-1404">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1404">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="c3a4a-1405">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1405">Object</span></span>|<span data-ttu-id="c3a4a-1406">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1406">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1407">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1407">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a4a-1408">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1408">Object</span></span>|<span data-ttu-id="c3a4a-1409">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1409">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1410">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1410">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c3a4a-1411">function</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1411">function</span></span>|<span data-ttu-id="c3a4a-1412">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1412">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1413">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1413">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c3a4a-1414">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1414">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c3a4a-1415">エラー</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1415">Errors</span></span>

|<span data-ttu-id="c3a4a-1416">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1416">Error code</span></span>|<span data-ttu-id="c3a4a-1417">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1417">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="c3a4a-1418">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1418">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a4a-1419">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1419">Requirements</span></span>

|<span data-ttu-id="c3a4a-1420">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1420">Requirement</span></span>|<span data-ttu-id="c3a4a-1421">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1421">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-1422">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1422">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-1423">1.1</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1423">1.1</span></span>|
|[<span data-ttu-id="c3a4a-1424">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1424">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-1425">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1425">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3a4a-1426">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1426">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-1427">作成</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1427">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-1428">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1428">Example</span></span>

<span data-ttu-id="c3a4a-1429">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1429">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="c3a4a-1430">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1430">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="c3a4a-1431">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1431">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="c3a4a-1432">現在、サポートされているイベントの種類は `Office.EventType.AttachmentsChanged`、`Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged`、`Office.EventType.RecurrenceChanged` です。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1432">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a4a-1433">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1433">Parameters</span></span>

| <span data-ttu-id="c3a4a-1434">名前</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1434">Name</span></span> | <span data-ttu-id="c3a4a-1435">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1435">Type</span></span> | <span data-ttu-id="c3a4a-1436">属性</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1436">Attributes</span></span> | <span data-ttu-id="c3a4a-1437">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1437">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c3a4a-1438">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1438">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c3a4a-1439">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1439">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="c3a4a-1440">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1440">Object</span></span> | <span data-ttu-id="c3a4a-1441">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1441">&lt;optional&gt;</span></span> | <span data-ttu-id="c3a4a-1442">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1442">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c3a4a-1443">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1443">Object</span></span> | <span data-ttu-id="c3a4a-1444">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1444">&lt;optional&gt;</span></span> | <span data-ttu-id="c3a4a-1445">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1445">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c3a4a-1446">function</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1446">function</span></span>| <span data-ttu-id="c3a4a-1447">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1447">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1448">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1448">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a4a-1449">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1449">Requirements</span></span>

|<span data-ttu-id="c3a4a-1450">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1450">Requirement</span></span>| <span data-ttu-id="c3a4a-1451">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1451">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-1452">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1452">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3a4a-1453">1.7</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1453">1.7</span></span> |
|[<span data-ttu-id="c3a4a-1454">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1454">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3a4a-1455">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1455">ReadItem</span></span> |
|[<span data-ttu-id="c3a4a-1456">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1456">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c3a4a-1457">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1457">Compose or Read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="c3a4a-1458">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1458">saveAsync([options], callback)</span></span>

<span data-ttu-id="c3a4a-1459">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1459">Asynchronously saves an item.</span></span>

<span data-ttu-id="c3a4a-p183">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p183">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a4a-1463">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1463">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="c3a4a-1464">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1464">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="c3a4a-p185">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p185">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a4a-1468">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1468">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="c3a4a-1469">Mac Outlook では、新規作成モードの会議で `saveAsync` をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1469">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="c3a4a-1470">Mac Outlook では、会議で `saveAsync` を呼び出すとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1470">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="c3a4a-1471">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1471">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a4a-1472">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1472">Parameters</span></span>

|<span data-ttu-id="c3a4a-1473">名前</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1473">Name</span></span>|<span data-ttu-id="c3a4a-1474">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1474">Type</span></span>|<span data-ttu-id="c3a4a-1475">属性</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1475">Attributes</span></span>|<span data-ttu-id="c3a4a-1476">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1476">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c3a4a-1477">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1477">Object</span></span>|<span data-ttu-id="c3a4a-1478">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1478">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1479">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1479">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a4a-1480">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1480">Object</span></span>|<span data-ttu-id="c3a4a-1481">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1481">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1482">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1482">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c3a4a-1483">function</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1483">function</span></span>||<span data-ttu-id="c3a4a-1484">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1484">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c3a4a-1485">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1485">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a4a-1486">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1486">Requirements</span></span>

|<span data-ttu-id="c3a4a-1487">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1487">Requirement</span></span>|<span data-ttu-id="c3a4a-1488">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1488">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-1489">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-1490">1.3</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1490">1.3</span></span>|
|[<span data-ttu-id="c3a4a-1491">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1491">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-1492">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1492">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3a4a-1493">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1493">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-1494">新規作成</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1494">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c3a4a-1495">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1495">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="c3a4a-p187">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p187">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="c3a4a-1498">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1498">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="c3a4a-1499">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1499">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="c3a4a-p188">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p188">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a4a-1503">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1503">Parameters</span></span>

|<span data-ttu-id="c3a4a-1504">名前</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1504">Name</span></span>|<span data-ttu-id="c3a4a-1505">型</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1505">Type</span></span>|<span data-ttu-id="c3a4a-1506">属性</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1506">Attributes</span></span>|<span data-ttu-id="c3a4a-1507">説明</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1507">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="c3a4a-1508">String</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1508">String</span></span>||<span data-ttu-id="c3a4a-p189">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p189">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="c3a4a-1512">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1512">Object</span></span>|<span data-ttu-id="c3a4a-1513">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1513">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1514">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1514">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a4a-1515">Object</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1515">Object</span></span>|<span data-ttu-id="c3a4a-1516">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1516">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-1517">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1517">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="c3a4a-1518">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1518">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="c3a4a-1519">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1519">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a4a-p190">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p190">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="c3a4a-p191">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-p191">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="c3a4a-1524">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1524">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="c3a4a-1525">function</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1525">function</span></span>||<span data-ttu-id="c3a4a-1526">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1526">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a4a-1527">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1527">Requirements</span></span>

|<span data-ttu-id="c3a4a-1528">要件</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1528">Requirement</span></span>|<span data-ttu-id="c3a4a-1529">値</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1529">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a4a-1530">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1530">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a4a-1531">1.2</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1531">1.2</span></span>|
|[<span data-ttu-id="c3a4a-1532">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1532">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a4a-1533">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1533">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3a4a-1534">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1534">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a4a-1535">作成</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1535">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a4a-1536">例</span><span class="sxs-lookup"><span data-stu-id="c3a4a-1536">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
