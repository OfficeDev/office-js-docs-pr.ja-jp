---
title: Office. アイテム-プレビュー要件セット
description: ''
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: 2ebcacb1f99df047b5f5c5ebe82c012e21e45d3c
ms.sourcegitcommit: 44f1a4a3e1ae3c33d7d5fabcee14b84af94e03da
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/03/2019
ms.locfileid: "39670140"
---
# <a name="item"></a><span data-ttu-id="c3a7b-102">item</span><span class="sxs-lookup"><span data-stu-id="c3a7b-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="c3a7b-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="c3a7b-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="c3a7b-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-mailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-mailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-106">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-106">Requirements</span></span>

|<span data-ttu-id="c3a7b-107">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-107">Requirement</span></span>|<span data-ttu-id="c3a7b-108">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-110">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-110">1.0</span></span>|
|[<span data-ttu-id="c3a7b-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="c3a7b-112">Restricted</span></span>|
|[<span data-ttu-id="c3a7b-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-114">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="c3a7b-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="c3a7b-115">Properties</span></span>

| <span data-ttu-id="c3a7b-116">プロパティ</span><span class="sxs-lookup"><span data-stu-id="c3a7b-116">Property</span></span> | <span data-ttu-id="c3a7b-117">最小値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-117">Minimum</span></span><br><span data-ttu-id="c3a7b-118">アクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-118">permission level</span></span> | <span data-ttu-id="c3a7b-119">モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-119">Modes</span></span> | <span data-ttu-id="c3a7b-120">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="c3a7b-120">Return type</span></span> | <span data-ttu-id="c3a7b-121">最小値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-121">Minimum</span></span><br><span data-ttu-id="c3a7b-122">要件セット</span><span class="sxs-lookup"><span data-stu-id="c3a7b-122">requirement set</span></span> |
|---|---|---|---|---|
| [<span data-ttu-id="c3a7b-123">attachments</span><span class="sxs-lookup"><span data-stu-id="c3a7b-123">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="c3a7b-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-124">ReadItem</span></span> | <span data-ttu-id="c3a7b-125">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-125">Read</span></span> | <span data-ttu-id="c3a7b-126">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c3a7b-126">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span> | <span data-ttu-id="c3a7b-127">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-127">1.0</span></span> |
| [<span data-ttu-id="c3a7b-128">bcc</span><span class="sxs-lookup"><span data-stu-id="c3a7b-128">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="c3a7b-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-129">ReadItem</span></span> | <span data-ttu-id="c3a7b-130">メッセージの作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-130">Message Compose</span></span> | [<span data-ttu-id="c3a7b-131">受信者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-131">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="c3a7b-132">1.1</span><span class="sxs-lookup"><span data-stu-id="c3a7b-132">1.1</span></span> |
| [<span data-ttu-id="c3a7b-133">body</span><span class="sxs-lookup"><span data-stu-id="c3a7b-133">body</span></span>](#body-body) | <span data-ttu-id="c3a7b-134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-134">ReadItem</span></span> | <span data-ttu-id="c3a7b-135">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-135">Compose</span></span> | [<span data-ttu-id="c3a7b-136">Body</span><span class="sxs-lookup"><span data-stu-id="c3a7b-136">Body</span></span>](/javascript/api/outlook/office.body) | <span data-ttu-id="c3a7b-137">1.1</span><span class="sxs-lookup"><span data-stu-id="c3a7b-137">1.1</span></span> |
| | | <span data-ttu-id="c3a7b-138">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-138">Read</span></span> | | |
| [<span data-ttu-id="c3a7b-139">categories</span><span class="sxs-lookup"><span data-stu-id="c3a7b-139">categories</span></span>](#categories-categories) | <span data-ttu-id="c3a7b-140">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-140">ReadItem</span></span> | <span data-ttu-id="c3a7b-141">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-141">Compose</span></span> | [<span data-ttu-id="c3a7b-142">Categories</span><span class="sxs-lookup"><span data-stu-id="c3a7b-142">Categories</span></span>](/javascript/api/outlook/office.categories) | <span data-ttu-id="c3a7b-143">1.8</span><span class="sxs-lookup"><span data-stu-id="c3a7b-143">1.8</span></span> |
| | | <span data-ttu-id="c3a7b-144">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-144">Read</span></span> | | |
| [<span data-ttu-id="c3a7b-145">cc</span><span class="sxs-lookup"><span data-stu-id="c3a7b-145">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="c3a7b-146">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-146">ReadItem</span></span> | <span data-ttu-id="c3a7b-147">メッセージの作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-147">Message Compose</span></span> | [<span data-ttu-id="c3a7b-148">受信者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-148">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="c3a7b-149">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-149">1.0</span></span> |
| | | <span data-ttu-id="c3a7b-150">メッセージの読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-150">Message Read</span></span> | <span data-ttu-id="c3a7b-151"><[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)></span><span class="sxs-lookup"><span data-stu-id="c3a7b-151">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span></span> | |
| [<span data-ttu-id="c3a7b-152">conversationId</span><span class="sxs-lookup"><span data-stu-id="c3a7b-152">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="c3a7b-153">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-153">ReadItem</span></span> | <span data-ttu-id="c3a7b-154">メッセージの作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-154">Message Compose</span></span> | <span data-ttu-id="c3a7b-155">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-155">String</span></span> | <span data-ttu-id="c3a7b-156">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-156">1.0</span></span> |
| | | <span data-ttu-id="c3a7b-157">メッセージの読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-157">Message Read</span></span> | | |
| [<span data-ttu-id="c3a7b-158">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="c3a7b-158">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="c3a7b-159">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-159">ReadItem</span></span> | <span data-ttu-id="c3a7b-160">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-160">Read</span></span> | <span data-ttu-id="c3a7b-161">日付</span><span class="sxs-lookup"><span data-stu-id="c3a7b-161">Date</span></span> | <span data-ttu-id="c3a7b-162">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-162">1.0</span></span> |
| [<span data-ttu-id="c3a7b-163">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="c3a7b-163">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="c3a7b-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-164">ReadItem</span></span> | <span data-ttu-id="c3a7b-165">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-165">Read</span></span> | <span data-ttu-id="c3a7b-166">日付</span><span class="sxs-lookup"><span data-stu-id="c3a7b-166">Date</span></span> | <span data-ttu-id="c3a7b-167">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-167">1.0</span></span> |
| [<span data-ttu-id="c3a7b-168">end</span><span class="sxs-lookup"><span data-stu-id="c3a7b-168">end</span></span>](#end-datetime) | <span data-ttu-id="c3a7b-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-169">ReadItem</span></span> | <span data-ttu-id="c3a7b-170">予定の開催者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-170">Appointment Organizer</span></span> | [<span data-ttu-id="c3a7b-171">Time</span><span class="sxs-lookup"><span data-stu-id="c3a7b-171">Time</span></span>](/javascript/api/outlook/office.time) | <span data-ttu-id="c3a7b-172">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-172">1.0</span></span> |
| | | <span data-ttu-id="c3a7b-173">予定の出席者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-173">Appointment Attendee</span></span> | <span data-ttu-id="c3a7b-174">日付</span><span class="sxs-lookup"><span data-stu-id="c3a7b-174">Date</span></span> | |
| | | <span data-ttu-id="c3a7b-175">メッセージの読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-175">Message Read</span></span><br><span data-ttu-id="c3a7b-176">(会議出席依頼)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-176">(Meeting Request)</span></span> | <span data-ttu-id="c3a7b-177">日付</span><span class="sxs-lookup"><span data-stu-id="c3a7b-177">Date</span></span> | |
| [<span data-ttu-id="c3a7b-178">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="c3a7b-178">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="c3a7b-179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-179">ReadItem</span></span> | <span data-ttu-id="c3a7b-180">予定の開催者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-180">Appointment Organizer</span></span> | [<span data-ttu-id="c3a7b-181">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="c3a7b-181">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation) | <span data-ttu-id="c3a7b-182">1.8</span><span class="sxs-lookup"><span data-stu-id="c3a7b-182">1.8</span></span> |
| | | <span data-ttu-id="c3a7b-183">予定の出席者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-183">Appointment Attendee</span></span> | | |
| [<span data-ttu-id="c3a7b-184">from</span><span class="sxs-lookup"><span data-stu-id="c3a7b-184">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="c3a7b-185">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-185">ReadWriteItem</span></span> | <span data-ttu-id="c3a7b-186">メッセージの作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-186">Message Compose</span></span> | [<span data-ttu-id="c3a7b-187">From</span><span class="sxs-lookup"><span data-stu-id="c3a7b-187">From</span></span>](/javascript/api/outlook/office.from) | <span data-ttu-id="c3a7b-188">1.7</span><span class="sxs-lookup"><span data-stu-id="c3a7b-188">1.7</span></span> |
| | <span data-ttu-id="c3a7b-189">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-189">ReadItem</span></span> | <span data-ttu-id="c3a7b-190">メッセージの読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-190">Message Read</span></span> | [<span data-ttu-id="c3a7b-191">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c3a7b-191">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails) | <span data-ttu-id="c3a7b-192">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-192">1.0</span></span> |
| [<span data-ttu-id="c3a7b-193">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="c3a7b-193">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="c3a7b-194">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-194">ReadItem</span></span> | <span data-ttu-id="c3a7b-195">メッセージの作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-195">Message Compose</span></span> | [<span data-ttu-id="c3a7b-196">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="c3a7b-196">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders) | <span data-ttu-id="c3a7b-197">1.8</span><span class="sxs-lookup"><span data-stu-id="c3a7b-197">1.8</span></span> |
| [<span data-ttu-id="c3a7b-198">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="c3a7b-198">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="c3a7b-199">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-199">ReadItem</span></span> | <span data-ttu-id="c3a7b-200">メッセージの読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-200">Message Read</span></span> | <span data-ttu-id="c3a7b-201">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-201">String</span></span> | <span data-ttu-id="c3a7b-202">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-202">1.0</span></span> |
| [<span data-ttu-id="c3a7b-203">itemClass</span><span class="sxs-lookup"><span data-stu-id="c3a7b-203">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="c3a7b-204">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-204">ReadItem</span></span> | <span data-ttu-id="c3a7b-205">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-205">Read</span></span> | <span data-ttu-id="c3a7b-206">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-206">String</span></span> | <span data-ttu-id="c3a7b-207">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-207">1.0</span></span> |
| [<span data-ttu-id="c3a7b-208">itemId</span><span class="sxs-lookup"><span data-stu-id="c3a7b-208">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="c3a7b-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-209">ReadItem</span></span> | <span data-ttu-id="c3a7b-210">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-210">Read</span></span> | <span data-ttu-id="c3a7b-211">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-211">String</span></span> | <span data-ttu-id="c3a7b-212">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-212">1.0</span></span> |
| [<span data-ttu-id="c3a7b-213">itemType</span><span class="sxs-lookup"><span data-stu-id="c3a7b-213">itemType</span></span>](#itemtype-mailboxenumsitemtype) | <span data-ttu-id="c3a7b-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-214">ReadItem</span></span> | <span data-ttu-id="c3a7b-215">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-215">Compose</span></span> | [<span data-ttu-id="c3a7b-216">MailboxEnums</span><span class="sxs-lookup"><span data-stu-id="c3a7b-216">MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype) | <span data-ttu-id="c3a7b-217">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-217">1.0</span></span> |
| | | <span data-ttu-id="c3a7b-218">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-218">Read</span></span> | | |
| [<span data-ttu-id="c3a7b-219">location</span><span class="sxs-lookup"><span data-stu-id="c3a7b-219">location</span></span>](#location-stringlocation) | <span data-ttu-id="c3a7b-220">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-220">ReadItem</span></span> | <span data-ttu-id="c3a7b-221">予定の開催者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-221">Appointment Organizer</span></span> | [<span data-ttu-id="c3a7b-222">Location</span><span class="sxs-lookup"><span data-stu-id="c3a7b-222">Location</span></span>](/javascript/api/outlook/office.location) | <span data-ttu-id="c3a7b-223">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-223">1.0</span></span> |
| | | <span data-ttu-id="c3a7b-224">予定の出席者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-224">Appointment Attendee</span></span> | <span data-ttu-id="c3a7b-225">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-225">String</span></span> | |
| | | <span data-ttu-id="c3a7b-226">メッセージの読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-226">Message Read</span></span><br><span data-ttu-id="c3a7b-227">(会議出席依頼)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-227">(Meeting Request)</span></span> | <span data-ttu-id="c3a7b-228">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-228">String</span></span> | |
| [<span data-ttu-id="c3a7b-229">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="c3a7b-229">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="c3a7b-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-230">ReadItem</span></span> | <span data-ttu-id="c3a7b-231">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-231">Read</span></span> | <span data-ttu-id="c3a7b-232">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-232">String</span></span> | <span data-ttu-id="c3a7b-233">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-233">1.0</span></span> |
| [<span data-ttu-id="c3a7b-234">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="c3a7b-234">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="c3a7b-235">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-235">ReadItem</span></span> | <span data-ttu-id="c3a7b-236">メッセージの作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-236">Message Compose</span></span> | [<span data-ttu-id="c3a7b-237">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="c3a7b-237">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages) | <span data-ttu-id="c3a7b-238">1.3</span><span class="sxs-lookup"><span data-stu-id="c3a7b-238">1.3</span></span> |
| | <span data-ttu-id="c3a7b-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-239">ReadItem</span></span> | <span data-ttu-id="c3a7b-240">メッセージの読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-240">Message Read</span></span> | | |
| [<span data-ttu-id="c3a7b-241">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="c3a7b-241">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="c3a7b-242">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-242">ReadItem</span></span> | <span data-ttu-id="c3a7b-243">予定の開催者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-243">Appointment Organizer</span></span> | [<span data-ttu-id="c3a7b-244">受信者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-244">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="c3a7b-245">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-245">1.0</span></span> |
| | | <span data-ttu-id="c3a7b-246">予定の出席者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-246">Appointment Attendee</span></span> | <span data-ttu-id="c3a7b-247"><[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)></span><span class="sxs-lookup"><span data-stu-id="c3a7b-247">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span></span> | |
| [<span data-ttu-id="c3a7b-248">organizer</span><span class="sxs-lookup"><span data-stu-id="c3a7b-248">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="c3a7b-249">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-249">ReadWriteItem</span></span> | <span data-ttu-id="c3a7b-250">予定の開催者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-250">Appointment Organizer</span></span> | [<span data-ttu-id="c3a7b-251">Organizer</span><span class="sxs-lookup"><span data-stu-id="c3a7b-251">Organizer</span></span>](/javascript/api/outlook/office.organizer) | <span data-ttu-id="c3a7b-252">1.7</span><span class="sxs-lookup"><span data-stu-id="c3a7b-252">1.7</span></span> |
| | <span data-ttu-id="c3a7b-253">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-253">ReadItem</span></span> | <span data-ttu-id="c3a7b-254">予定の出席者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-254">Appointment Attendee</span></span> | [<span data-ttu-id="c3a7b-255">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c3a7b-255">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails) | <span data-ttu-id="c3a7b-256">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-256">1.0</span></span> |
| [<span data-ttu-id="c3a7b-257">recurrence</span><span class="sxs-lookup"><span data-stu-id="c3a7b-257">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="c3a7b-258">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-258">ReadItem</span></span> | <span data-ttu-id="c3a7b-259">予定の開催者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-259">Appointment Organizer</span></span> | [<span data-ttu-id="c3a7b-260">繰り返さ</span><span class="sxs-lookup"><span data-stu-id="c3a7b-260">Recurrence</span></span>](/javascript/api/outlook/office.recurrence) | <span data-ttu-id="c3a7b-261">1.7</span><span class="sxs-lookup"><span data-stu-id="c3a7b-261">1.7</span></span> |
| | | <span data-ttu-id="c3a7b-262">予定の出席者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-262">Appointment Attendee</span></span> | | |
| | | <span data-ttu-id="c3a7b-263">メッセージの読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-263">Message Read</span></span><br><span data-ttu-id="c3a7b-264">(会議出席依頼)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-264">(Meeting Request)</span></span> | | |
| [<span data-ttu-id="c3a7b-265">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="c3a7b-265">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="c3a7b-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-266">ReadItem</span></span> | <span data-ttu-id="c3a7b-267">予定の開催者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-267">Appointment Organizer</span></span> | [<span data-ttu-id="c3a7b-268">受信者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-268">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="c3a7b-269">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-269">1.0</span></span> |
| | | <span data-ttu-id="c3a7b-270">予定の出席者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-270">Appointment Attendee</span></span> | <span data-ttu-id="c3a7b-271"><[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)></span><span class="sxs-lookup"><span data-stu-id="c3a7b-271">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span></span> | |
| [<span data-ttu-id="c3a7b-272">sender</span><span class="sxs-lookup"><span data-stu-id="c3a7b-272">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="c3a7b-273">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-273">ReadItem</span></span> | <span data-ttu-id="c3a7b-274">メッセージの読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-274">Message Read</span></span> | [<span data-ttu-id="c3a7b-275">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c3a7b-275">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails) | <span data-ttu-id="c3a7b-276">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-276">1.0</span></span> |
| [<span data-ttu-id="c3a7b-277">系列 Id</span><span class="sxs-lookup"><span data-stu-id="c3a7b-277">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="c3a7b-278">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-278">ReadItem</span></span> | <span data-ttu-id="c3a7b-279">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-279">Compose</span></span> | <span data-ttu-id="c3a7b-280">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-280">String</span></span> | <span data-ttu-id="c3a7b-281">1.7</span><span class="sxs-lookup"><span data-stu-id="c3a7b-281">1.7</span></span> |
| | | <span data-ttu-id="c3a7b-282">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-282">Read</span></span> | | |
| [<span data-ttu-id="c3a7b-283">start</span><span class="sxs-lookup"><span data-stu-id="c3a7b-283">start</span></span>](#start-datetime) | <span data-ttu-id="c3a7b-284">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-284">ReadItem</span></span> | <span data-ttu-id="c3a7b-285">予定の開催者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-285">Appointment Organizer</span></span> | [<span data-ttu-id="c3a7b-286">Time</span><span class="sxs-lookup"><span data-stu-id="c3a7b-286">Time</span></span>](/javascript/api/outlook/office.time) | <span data-ttu-id="c3a7b-287">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-287">1.0</span></span> |
| | | <span data-ttu-id="c3a7b-288">予定の出席者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-288">Appointment Attendee</span></span> | <span data-ttu-id="c3a7b-289">日付</span><span class="sxs-lookup"><span data-stu-id="c3a7b-289">Date</span></span> | |
| | | <span data-ttu-id="c3a7b-290">メッセージの読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-290">Message Read</span></span><br><span data-ttu-id="c3a7b-291">(会議出席依頼)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-291">(Meeting Request)</span></span> | <span data-ttu-id="c3a7b-292">日付</span><span class="sxs-lookup"><span data-stu-id="c3a7b-292">Date</span></span> | |
| [<span data-ttu-id="c3a7b-293">subject</span><span class="sxs-lookup"><span data-stu-id="c3a7b-293">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="c3a7b-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-294">ReadItem</span></span> | <span data-ttu-id="c3a7b-295">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-295">Compose</span></span> | [<span data-ttu-id="c3a7b-296">件名</span><span class="sxs-lookup"><span data-stu-id="c3a7b-296">Subject</span></span>](/javascript/api/outlook/office.subject) | <span data-ttu-id="c3a7b-297">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-297">1.0</span></span> |
| | | <span data-ttu-id="c3a7b-298">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-298">Read</span></span> | <span data-ttu-id="c3a7b-299">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-299">String</span></span> | |
| [<span data-ttu-id="c3a7b-300">to</span><span class="sxs-lookup"><span data-stu-id="c3a7b-300">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="c3a7b-301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-301">ReadItem</span></span> | <span data-ttu-id="c3a7b-302">メッセージの作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-302">Message Compose</span></span> | [<span data-ttu-id="c3a7b-303">受信者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-303">Recipients</span></span>](/javascript/api/outlook/office.recipients) | <span data-ttu-id="c3a7b-304">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-304">1.0</span></span> |
| | | <span data-ttu-id="c3a7b-305">メッセージの読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-305">Message Read</span></span> | <span data-ttu-id="c3a7b-306"><[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)></span><span class="sxs-lookup"><span data-stu-id="c3a7b-306">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)></span></span> | |

##### <a name="methods"></a><span data-ttu-id="c3a7b-307">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a7b-307">Methods</span></span>

| <span data-ttu-id="c3a7b-308">メソッド</span><span class="sxs-lookup"><span data-stu-id="c3a7b-308">Method</span></span> | <span data-ttu-id="c3a7b-309">最小値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-309">Minimum</span></span><br><span data-ttu-id="c3a7b-310">アクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-310">permission level</span></span> | <span data-ttu-id="c3a7b-311">モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-311">Modes</span></span> | <span data-ttu-id="c3a7b-312">最小値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-312">Minimum</span></span><br><span data-ttu-id="c3a7b-313">要件セット</span><span class="sxs-lookup"><span data-stu-id="c3a7b-313">requirement set</span></span> |
|---|---|---|---|
| [<span data-ttu-id="c3a7b-314">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c3a7b-314">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="c3a7b-315">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-315">ReadWriteItem</span></span> | <span data-ttu-id="c3a7b-316">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-316">Compose</span></span> | <span data-ttu-id="c3a7b-317">1.1</span><span class="sxs-lookup"><span data-stu-id="c3a7b-317">1.1</span></span> |
| [<span data-ttu-id="c3a7b-318">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="c3a7b-318">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="c3a7b-319">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-319">ReadWriteItem</span></span> | <span data-ttu-id="c3a7b-320">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-320">Compose</span></span> | <span data-ttu-id="c3a7b-321">1.8</span><span class="sxs-lookup"><span data-stu-id="c3a7b-321">1.8</span></span> |
| [<span data-ttu-id="c3a7b-322">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c3a7b-322">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="c3a7b-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-323">ReadItem</span></span> | <span data-ttu-id="c3a7b-324">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-324">Compose</span></span><br><span data-ttu-id="c3a7b-325">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-325">Read</span></span> | <span data-ttu-id="c3a7b-326">1.7</span><span class="sxs-lookup"><span data-stu-id="c3a7b-326">1.7</span></span> |
| [<span data-ttu-id="c3a7b-327">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c3a7b-327">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="c3a7b-328">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-328">ReadWriteItem</span></span> | <span data-ttu-id="c3a7b-329">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-329">Compose</span></span> | <span data-ttu-id="c3a7b-330">1.1</span><span class="sxs-lookup"><span data-stu-id="c3a7b-330">1.1</span></span> |
| [<span data-ttu-id="c3a7b-331">close</span><span class="sxs-lookup"><span data-stu-id="c3a7b-331">close</span></span>](#close) | <span data-ttu-id="c3a7b-332">制限あり</span><span class="sxs-lookup"><span data-stu-id="c3a7b-332">Restricted</span></span> | <span data-ttu-id="c3a7b-333">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-333">Compose</span></span> | <span data-ttu-id="c3a7b-334">1.3</span><span class="sxs-lookup"><span data-stu-id="c3a7b-334">1.3</span></span> |
| [<span data-ttu-id="c3a7b-335">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="c3a7b-335">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="c3a7b-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-336">ReadItem</span></span> | <span data-ttu-id="c3a7b-337">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-337">Read</span></span> | <span data-ttu-id="c3a7b-338">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-338">1.0</span></span> |
| [<span data-ttu-id="c3a7b-339">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="c3a7b-339">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="c3a7b-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-340">ReadItem</span></span> | <span data-ttu-id="c3a7b-341">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-341">Read</span></span> | <span data-ttu-id="c3a7b-342">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-342">1.0</span></span> |
| [<span data-ttu-id="c3a7b-343">getAllInternetHeadersAsync</span><span class="sxs-lookup"><span data-stu-id="c3a7b-343">getAllInternetHeadersAsync</span></span>](#getallinternetheadersasyncoptions-callback) | <span data-ttu-id="c3a7b-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-344">ReadItem</span></span> | <span data-ttu-id="c3a7b-345">メッセージの読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-345">Message Read</span></span> | <span data-ttu-id="c3a7b-346">1.8</span><span class="sxs-lookup"><span data-stu-id="c3a7b-346">1.8</span></span> |
| [<span data-ttu-id="c3a7b-347">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="c3a7b-347">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="c3a7b-348">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-348">ReadItem</span></span> | <span data-ttu-id="c3a7b-349">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-349">Compose</span></span><br><span data-ttu-id="c3a7b-350">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-350">Read</span></span> | <span data-ttu-id="c3a7b-351">1.8</span><span class="sxs-lookup"><span data-stu-id="c3a7b-351">1.8</span></span> |
| [<span data-ttu-id="c3a7b-352">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="c3a7b-352">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="c3a7b-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-353">ReadItem</span></span> | <span data-ttu-id="c3a7b-354">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-354">Compose</span></span> | <span data-ttu-id="c3a7b-355">1.8</span><span class="sxs-lookup"><span data-stu-id="c3a7b-355">1.8</span></span> |
| [<span data-ttu-id="c3a7b-356">getEntities</span><span class="sxs-lookup"><span data-stu-id="c3a7b-356">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="c3a7b-357">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-357">ReadItem</span></span> | <span data-ttu-id="c3a7b-358">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-358">Read</span></span> | <span data-ttu-id="c3a7b-359">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-359">1.0</span></span> |
| [<span data-ttu-id="c3a7b-360">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="c3a7b-360">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="c3a7b-361">制限あり</span><span class="sxs-lookup"><span data-stu-id="c3a7b-361">Restricted</span></span> | <span data-ttu-id="c3a7b-362">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-362">Read</span></span> | <span data-ttu-id="c3a7b-363">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-363">1.0</span></span> |
| [<span data-ttu-id="c3a7b-364">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="c3a7b-364">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="c3a7b-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-365">ReadItem</span></span> | <span data-ttu-id="c3a7b-366">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-366">Read</span></span> | <span data-ttu-id="c3a7b-367">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-367">1.0</span></span> |
| [<span data-ttu-id="c3a7b-368">、Office.context.mailbox.item.getinitializationcontextasync</span><span class="sxs-lookup"><span data-stu-id="c3a7b-368">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="c3a7b-369">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-369">ReadItem</span></span> | <span data-ttu-id="c3a7b-370">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-370">Read</span></span> | <span data-ttu-id="c3a7b-371">プレビュー</span><span class="sxs-lookup"><span data-stu-id="c3a7b-371">Preview</span></span> |
| [<span data-ttu-id="c3a7b-372">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="c3a7b-372">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="c3a7b-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-373">ReadItem</span></span> | <span data-ttu-id="c3a7b-374">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-374">Compose</span></span> | <span data-ttu-id="c3a7b-375">1.8</span><span class="sxs-lookup"><span data-stu-id="c3a7b-375">1.8</span></span> |
| [<span data-ttu-id="c3a7b-376">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="c3a7b-376">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="c3a7b-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-377">ReadItem</span></span> | <span data-ttu-id="c3a7b-378">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-378">Read</span></span> | <span data-ttu-id="c3a7b-379">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-379">1.0</span></span> |
| [<span data-ttu-id="c3a7b-380">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="c3a7b-380">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="c3a7b-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-381">ReadItem</span></span> | <span data-ttu-id="c3a7b-382">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-382">Read</span></span> | <span data-ttu-id="c3a7b-383">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-383">1.0</span></span> |
| [<span data-ttu-id="c3a7b-384">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c3a7b-384">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="c3a7b-385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-385">ReadItem</span></span> | <span data-ttu-id="c3a7b-386">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-386">Compose</span></span> | <span data-ttu-id="c3a7b-387">1.2</span><span class="sxs-lookup"><span data-stu-id="c3a7b-387">1.2</span></span> |
| [<span data-ttu-id="c3a7b-388">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="c3a7b-388">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="c3a7b-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-389">ReadItem</span></span> | <span data-ttu-id="c3a7b-390">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-390">Read</span></span> | <span data-ttu-id="c3a7b-391">1.6</span><span class="sxs-lookup"><span data-stu-id="c3a7b-391">1.6</span></span> |
| [<span data-ttu-id="c3a7b-392">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="c3a7b-392">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="c3a7b-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-393">ReadItem</span></span> | <span data-ttu-id="c3a7b-394">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-394">Read</span></span> | <span data-ttu-id="c3a7b-395">1.6</span><span class="sxs-lookup"><span data-stu-id="c3a7b-395">1.6</span></span> |
| [<span data-ttu-id="c3a7b-396">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="c3a7b-396">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="c3a7b-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-397">ReadItem</span></span> | <span data-ttu-id="c3a7b-398">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-398">Compose</span></span><br><span data-ttu-id="c3a7b-399">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-399">Read</span></span> | <span data-ttu-id="c3a7b-400">1.8</span><span class="sxs-lookup"><span data-stu-id="c3a7b-400">1.8</span></span> |
| [<span data-ttu-id="c3a7b-401">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="c3a7b-401">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="c3a7b-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-402">ReadItem</span></span> | <span data-ttu-id="c3a7b-403">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-403">Compose</span></span><br><span data-ttu-id="c3a7b-404">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-404">Read</span></span> | <span data-ttu-id="c3a7b-405">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-405">1.0</span></span> |
| [<span data-ttu-id="c3a7b-406">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c3a7b-406">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="c3a7b-407">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-407">ReadWriteItem</span></span> | <span data-ttu-id="c3a7b-408">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-408">Compose</span></span> | <span data-ttu-id="c3a7b-409">1.1</span><span class="sxs-lookup"><span data-stu-id="c3a7b-409">1.1</span></span> |
| [<span data-ttu-id="c3a7b-410">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c3a7b-410">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="c3a7b-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-411">ReadItem</span></span> | <span data-ttu-id="c3a7b-412">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-412">Compose</span></span><br><span data-ttu-id="c3a7b-413">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-413">Read</span></span> | <span data-ttu-id="c3a7b-414">1.7</span><span class="sxs-lookup"><span data-stu-id="c3a7b-414">1.7</span></span> |
| [<span data-ttu-id="c3a7b-415">saveAsync</span><span class="sxs-lookup"><span data-stu-id="c3a7b-415">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="c3a7b-416">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-416">ReadWriteItem</span></span> | <span data-ttu-id="c3a7b-417">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-417">Compose</span></span> | <span data-ttu-id="c3a7b-418">1.3</span><span class="sxs-lookup"><span data-stu-id="c3a7b-418">1.3</span></span> |
| [<span data-ttu-id="c3a7b-419">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c3a7b-419">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="c3a7b-420">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-420">ReadWriteItem</span></span> | <span data-ttu-id="c3a7b-421">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-421">Compose</span></span> | <span data-ttu-id="c3a7b-422">1.2</span><span class="sxs-lookup"><span data-stu-id="c3a7b-422">1.2</span></span> |

##### <a name="events"></a><span data-ttu-id="c3a7b-423">イベント</span><span class="sxs-lookup"><span data-stu-id="c3a7b-423">Events</span></span>

<span data-ttu-id="c3a7b-424">[Addハンドラ async](#addhandlerasynceventtype-handler-options-callback)と[removeハンドラ async](#removehandlerasynceventtype-options-callback)を使用して、次のイベントにサブスクライブし、サブスクライブを解除することができます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-424">You can subscribe to and unsubscribe from the following events using [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) and [removeHandlerAsync](#removehandlerasynceventtype-options-callback) respectively.</span></span>

| <span data-ttu-id="c3a7b-425">イベント</span><span class="sxs-lookup"><span data-stu-id="c3a7b-425">Event</span></span> | <span data-ttu-id="c3a7b-426">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-426">Description</span></span> | <span data-ttu-id="c3a7b-427">最小値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-427">Minimum</span></span><br><span data-ttu-id="c3a7b-428">要件セット</span><span class="sxs-lookup"><span data-stu-id="c3a7b-428">requirement set</span></span> |
|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="c3a7b-429">選択した予定またはデータ系列の日付または時刻が変更されました。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-429">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="c3a7b-430">1.7</span><span class="sxs-lookup"><span data-stu-id="c3a7b-430">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="c3a7b-431">添付ファイルがアイテムに追加またはアイテムから削除されています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-431">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="c3a7b-432">1.8</span><span class="sxs-lookup"><span data-stu-id="c3a7b-432">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="c3a7b-433">選択した予定の場所が変更されました。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-433">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="c3a7b-434">1.8</span><span class="sxs-lookup"><span data-stu-id="c3a7b-434">1.8</span></span> |
|`RecipientsChanged`| <span data-ttu-id="c3a7b-435">選択したアイテムまたは予定の場所の受信者の一覧が変更されました。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-435">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="c3a7b-436">1.7</span><span class="sxs-lookup"><span data-stu-id="c3a7b-436">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="c3a7b-437">選択したアイテムの定期的なパターンが変更されました。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-437">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="c3a7b-438">1.7</span><span class="sxs-lookup"><span data-stu-id="c3a7b-438">1.7</span></span> |

### <a name="example"></a><span data-ttu-id="c3a7b-439">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-439">Example</span></span>

<span data-ttu-id="c3a7b-440">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-440">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

## <a name="property-details"></a><span data-ttu-id="c3a7b-441">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="c3a7b-441">Property details</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="c3a7b-442">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c3a7b-442">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="c3a7b-443">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-443">Gets the item's attachments as an array.</span></span> <span data-ttu-id="c3a7b-444">閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-444">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a7b-445">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-445">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="c3a7b-446">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-446">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="c3a7b-447">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-447">Type</span></span>

*   <span data-ttu-id="c3a7b-448">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c3a7b-448">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-449">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-449">Requirements</span></span>

|<span data-ttu-id="c3a7b-450">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-450">Requirement</span></span>|<span data-ttu-id="c3a7b-451">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-451">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-452">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-452">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-453">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-453">1.0</span></span>|
|[<span data-ttu-id="c3a7b-454">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-454">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-455">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-455">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-456">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-456">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-457">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-457">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-458">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-458">Example</span></span>

<span data-ttu-id="c3a7b-459">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-459">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="c3a7b-460">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-460">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="c3a7b-461">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-461">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="c3a7b-462">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-462">Compose mode only.</span></span>

<span data-ttu-id="c3a7b-463">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-463">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c3a7b-464">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-464">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="c3a7b-465">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-465">Get 500 members maximum.</span></span>
- <span data-ttu-id="c3a7b-466">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-466">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a7b-467">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-467">Type</span></span>

*   [<span data-ttu-id="c3a7b-468">受信者</span><span class="sxs-lookup"><span data-stu-id="c3a7b-468">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="c3a7b-469">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-469">Requirements</span></span>

|<span data-ttu-id="c3a7b-470">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-470">Requirement</span></span>|<span data-ttu-id="c3a7b-471">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-472">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-473">1.1</span><span class="sxs-lookup"><span data-stu-id="c3a7b-473">1.1</span></span>|
|[<span data-ttu-id="c3a7b-474">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-474">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-475">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-476">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-476">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-477">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-477">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-478">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-478">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="c3a7b-479">body: [Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-479">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="c3a7b-480">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-480">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a7b-481">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-481">Type</span></span>

*   [<span data-ttu-id="c3a7b-482">Body</span><span class="sxs-lookup"><span data-stu-id="c3a7b-482">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="c3a7b-483">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-483">Requirements</span></span>

|<span data-ttu-id="c3a7b-484">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-484">Requirement</span></span>|<span data-ttu-id="c3a7b-485">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-485">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-486">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-486">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-487">1.1</span><span class="sxs-lookup"><span data-stu-id="c3a7b-487">1.1</span></span>|
|[<span data-ttu-id="c3a7b-488">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-488">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-489">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-489">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-490">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-490">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-491">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-491">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-492">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-492">Example</span></span>

<span data-ttu-id="c3a7b-493">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-493">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="c3a7b-494">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-494">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="c3a7b-495">カテゴリ:[カテゴリ](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-495">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="c3a7b-496">アイテムのカテゴリを管理するためのメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-496">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a7b-497">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-497">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a7b-498">種類</span><span class="sxs-lookup"><span data-stu-id="c3a7b-498">Type</span></span>

*   [<span data-ttu-id="c3a7b-499">Categories</span><span class="sxs-lookup"><span data-stu-id="c3a7b-499">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="c3a7b-500">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-500">Requirements</span></span>

|<span data-ttu-id="c3a7b-501">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-501">Requirement</span></span>|<span data-ttu-id="c3a7b-502">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-503">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-504">1.8</span><span class="sxs-lookup"><span data-stu-id="c3a7b-504">1.8</span></span>|
|[<span data-ttu-id="c3a7b-505">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-506">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-507">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-508">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-508">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-509">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-509">Example</span></span>

<span data-ttu-id="c3a7b-510">この例では、アイテムのカテゴリを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-510">This example gets the item's categories.</span></span>

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="c3a7b-511">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-511">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="c3a7b-512">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-512">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="c3a7b-513">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-513">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a7b-514">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-514">Read mode</span></span>

<span data-ttu-id="c3a7b-515">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-515">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="c3a7b-516">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-516">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c3a7b-517">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-517">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a7b-518">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-518">Compose mode</span></span>

<span data-ttu-id="c3a7b-519">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-519">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="c3a7b-520">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-520">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c3a7b-521">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-521">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="c3a7b-522">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-522">Get 500 members maximum.</span></span>
- <span data-ttu-id="c3a7b-523">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-523">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c3a7b-524">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-524">Type</span></span>

*   <span data-ttu-id="c3a7b-525">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-525">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-526">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-526">Requirements</span></span>

|<span data-ttu-id="c3a7b-527">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-527">Requirement</span></span>|<span data-ttu-id="c3a7b-528">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-528">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-529">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-529">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-530">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-530">1.0</span></span>|
|[<span data-ttu-id="c3a7b-531">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-531">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-532">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-532">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-533">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-533">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-534">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-534">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="c3a7b-535">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-535">(nullable) conversationId: String</span></span>

<span data-ttu-id="c3a7b-536">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-536">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="c3a7b-p109">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="c3a7b-p110">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a7b-541">Type</span><span class="sxs-lookup"><span data-stu-id="c3a7b-541">Type</span></span>

*   <span data-ttu-id="c3a7b-542">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-542">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-543">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-543">Requirements</span></span>

|<span data-ttu-id="c3a7b-544">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-544">Requirement</span></span>|<span data-ttu-id="c3a7b-545">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-545">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-546">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-547">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-547">1.0</span></span>|
|[<span data-ttu-id="c3a7b-548">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-548">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-549">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-549">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-550">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-550">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-551">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-551">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-552">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-552">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="c3a7b-553">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="c3a7b-553">dateTimeCreated: Date</span></span>

<span data-ttu-id="c3a7b-p111">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a7b-556">種類</span><span class="sxs-lookup"><span data-stu-id="c3a7b-556">Type</span></span>

*   <span data-ttu-id="c3a7b-557">日付</span><span class="sxs-lookup"><span data-stu-id="c3a7b-557">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-558">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-558">Requirements</span></span>

|<span data-ttu-id="c3a7b-559">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-559">Requirement</span></span>|<span data-ttu-id="c3a7b-560">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-561">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-562">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-562">1.0</span></span>|
|[<span data-ttu-id="c3a7b-563">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-563">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-564">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-565">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-565">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-566">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-566">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-567">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-567">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="c3a7b-568">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="c3a7b-568">dateTimeModified: Date</span></span>

<span data-ttu-id="c3a7b-p112">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a7b-571">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-571">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a7b-572">種類</span><span class="sxs-lookup"><span data-stu-id="c3a7b-572">Type</span></span>

*   <span data-ttu-id="c3a7b-573">日付</span><span class="sxs-lookup"><span data-stu-id="c3a7b-573">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-574">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-574">Requirements</span></span>

|<span data-ttu-id="c3a7b-575">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-575">Requirement</span></span>|<span data-ttu-id="c3a7b-576">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-576">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-577">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-577">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-578">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-578">1.0</span></span>|
|[<span data-ttu-id="c3a7b-579">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-579">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-580">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-580">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-581">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-581">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-582">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-582">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-583">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-583">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="c3a7b-584">end: Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-584">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="c3a7b-585">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-585">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="c3a7b-p113">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a7b-588">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-588">Read mode</span></span>

<span data-ttu-id="c3a7b-589">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-589">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a7b-590">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-590">Compose mode</span></span>

<span data-ttu-id="c3a7b-591">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-591">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="c3a7b-592">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-592">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="c3a7b-593">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-593">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="c3a7b-594">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-594">Type</span></span>

*   <span data-ttu-id="c3a7b-595">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-595">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-596">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-596">Requirements</span></span>

|<span data-ttu-id="c3a7b-597">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-597">Requirement</span></span>|<span data-ttu-id="c3a7b-598">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-598">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-599">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-599">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-600">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-600">1.0</span></span>|
|[<span data-ttu-id="c3a7b-601">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-601">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-602">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-602">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-603">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-603">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-604">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-604">Compose or Read</span></span>|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="c3a7b-605">enhancedLocation: [enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-605">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="c3a7b-606">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-606">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a7b-607">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-607">Read mode</span></span>

<span data-ttu-id="c3a7b-608">この`enhancedLocation`プロパティは、予定に関連付けられている場所 ( [locationdetails](/javascript/api/outlook/office.locationdetails)オブジェクトで表される) のセットを取得できる[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-608">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c3a7b-609">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-609">Compose mode</span></span>

<span data-ttu-id="c3a7b-610">この`enhancedLocation`プロパティは、予定の場所を取得、削除、または追加するためのメソッドを提供する[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-610">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a7b-611">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-611">Type</span></span>

*   [<span data-ttu-id="c3a7b-612">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="c3a7b-612">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="c3a7b-613">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-613">Requirements</span></span>

|<span data-ttu-id="c3a7b-614">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-614">Requirement</span></span>|<span data-ttu-id="c3a7b-615">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-615">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-616">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-616">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-617">1.8</span><span class="sxs-lookup"><span data-stu-id="c3a7b-617">1.8</span></span>|
|[<span data-ttu-id="c3a7b-618">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-618">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-619">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-619">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-620">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-620">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-621">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-621">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-622">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-622">Example</span></span>

<span data-ttu-id="c3a7b-623">次の例では、予定に関連付けられている現在の場所を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-623">The following example gets the current locations associated with the appointment.</span></span>

```js
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

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="c3a7b-624">from: [emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)|[from](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-624">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="c3a7b-625">メッセージの送信者の電子メール アドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-625">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="c3a7b-p114">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a7b-628">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-628">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a7b-629">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-629">Read mode</span></span>

<span data-ttu-id="c3a7b-630">プロパティ`from`は`EmailAddressDetails`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-630">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a7b-631">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-631">Compose mode</span></span>

<span data-ttu-id="c3a7b-632">プロパティ`from`は、from `From`値を取得するメソッドを提供するオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-632">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c3a7b-633">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-633">Type</span></span>

*   <span data-ttu-id="c3a7b-634">[電子メールアドレス](/javascript/api/outlook/office.emailaddressdetails) | [の](/javascript/api/outlook/office.from)詳細</span><span class="sxs-lookup"><span data-stu-id="c3a7b-634">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-635">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-635">Requirements</span></span>

|<span data-ttu-id="c3a7b-636">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-636">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="c3a7b-637">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-637">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-638">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-638">1.0</span></span>|<span data-ttu-id="c3a7b-639">1.7</span><span class="sxs-lookup"><span data-stu-id="c3a7b-639">1.7</span></span>|
|[<span data-ttu-id="c3a7b-640">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-640">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-641">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-641">ReadItem</span></span>|<span data-ttu-id="c3a7b-642">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-642">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3a7b-643">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-643">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-644">Read</span><span class="sxs-lookup"><span data-stu-id="c3a7b-644">Read</span></span>|<span data-ttu-id="c3a7b-645">Compose</span><span class="sxs-lookup"><span data-stu-id="c3a7b-645">Compose</span></span>|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="c3a7b-646">internetHeaders: [internetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-646">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="c3a7b-647">メッセージのカスタムインターネットヘッダーを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-647">Gets or sets custom internet headers on a message.</span></span> <span data-ttu-id="c3a7b-648">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-648">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a7b-649">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-649">Type</span></span>

*   [<span data-ttu-id="c3a7b-650">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="c3a7b-650">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="c3a7b-651">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-651">Requirements</span></span>

|<span data-ttu-id="c3a7b-652">必要条件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-652">Requirement</span></span>|<span data-ttu-id="c3a7b-653">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-653">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-654">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-654">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-655">1.8</span><span class="sxs-lookup"><span data-stu-id="c3a7b-655">1.8</span></span>|
|[<span data-ttu-id="c3a7b-656">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-656">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-657">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-657">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-658">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-658">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-659">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-659">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-660">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-660">Example</span></span>

```js
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="c3a7b-661">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-661">internetMessageId: String</span></span>

<span data-ttu-id="c3a7b-p116">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a7b-664">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-664">Type</span></span>

*   <span data-ttu-id="c3a7b-665">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-665">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-666">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-666">Requirements</span></span>

|<span data-ttu-id="c3a7b-667">必要条件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-667">Requirement</span></span>|<span data-ttu-id="c3a7b-668">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-669">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-670">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-670">1.0</span></span>|
|[<span data-ttu-id="c3a7b-671">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-671">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-672">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-673">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-673">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-674">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-674">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-675">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-675">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="c3a7b-676">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-676">itemClass: String</span></span>

<span data-ttu-id="c3a7b-p117">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="c3a7b-p118">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="c3a7b-681">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-681">Type</span></span>|<span data-ttu-id="c3a7b-682">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-682">Description</span></span>|<span data-ttu-id="c3a7b-683">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="c3a7b-683">item class</span></span>|
|---|---|---|
|<span data-ttu-id="c3a7b-684">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="c3a7b-684">Appointment items</span></span>|<span data-ttu-id="c3a7b-685">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-685">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="c3a7b-686">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="c3a7b-686">Message items</span></span>|<span data-ttu-id="c3a7b-687">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-687">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="c3a7b-688">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-688">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a7b-689">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-689">Type</span></span>

*   <span data-ttu-id="c3a7b-690">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-690">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-691">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-691">Requirements</span></span>

|<span data-ttu-id="c3a7b-692">必要条件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-692">Requirement</span></span>|<span data-ttu-id="c3a7b-693">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-694">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-695">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-695">1.0</span></span>|
|[<span data-ttu-id="c3a7b-696">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-696">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-697">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-697">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-698">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-698">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-699">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-699">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-700">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-700">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="c3a7b-701">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-701">(nullable) itemId: String</span></span>

<span data-ttu-id="c3a7b-p119">現在のアイテムの [Exchange Web サービスのアイテム識別子](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p119">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a7b-704">`itemId` プロパティから返される識別子は、[Exchange Web サービスのアイテム識別子](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) と同じです。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-704">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="c3a7b-705">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-705">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="c3a7b-706">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-706">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="c3a7b-707">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-707">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="c3a7b-p121">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a7b-710">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-710">Type</span></span>

*   <span data-ttu-id="c3a7b-711">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-711">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-712">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-712">Requirements</span></span>

|<span data-ttu-id="c3a7b-713">必要条件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-713">Requirement</span></span>|<span data-ttu-id="c3a7b-714">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-714">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-715">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-715">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-716">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-716">1.0</span></span>|
|[<span data-ttu-id="c3a7b-717">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-717">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-718">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-718">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-719">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-719">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-720">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-720">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-721">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-721">Example</span></span>

<span data-ttu-id="c3a7b-p122">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-mailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="c3a7b-724">itemType: [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-724">itemType: [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="c3a7b-725">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-725">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="c3a7b-726">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-726">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a7b-727">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-727">Type</span></span>

*   [<span data-ttu-id="c3a7b-728">MailboxEnums</span><span class="sxs-lookup"><span data-stu-id="c3a7b-728">MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="c3a7b-729">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-729">Requirements</span></span>

|<span data-ttu-id="c3a7b-730">必要条件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-730">Requirement</span></span>|<span data-ttu-id="c3a7b-731">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-731">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-732">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-732">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-733">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-733">1.0</span></span>|
|[<span data-ttu-id="c3a7b-734">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-734">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-735">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-735">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-736">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-736">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-737">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-737">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-738">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-738">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="c3a7b-739">location: String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-739">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="c3a7b-740">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-740">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a7b-741">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-741">Read mode</span></span>

<span data-ttu-id="c3a7b-742">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-742">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a7b-743">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-743">Compose mode</span></span>

<span data-ttu-id="c3a7b-744">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-744">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c3a7b-745">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-745">Type</span></span>

*   <span data-ttu-id="c3a7b-746">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-746">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-747">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-747">Requirements</span></span>

|<span data-ttu-id="c3a7b-748">必要条件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-748">Requirement</span></span>|<span data-ttu-id="c3a7b-749">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-749">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-750">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-750">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-751">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-751">1.0</span></span>|
|[<span data-ttu-id="c3a7b-752">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-752">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-753">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-753">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-754">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-754">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-755">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-755">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="c3a7b-756">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-756">normalizedSubject: String</span></span>

<span data-ttu-id="c3a7b-p123">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="c3a7b-p124">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a7b-761">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-761">Type</span></span>

*   <span data-ttu-id="c3a7b-762">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-762">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-763">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-763">Requirements</span></span>

|<span data-ttu-id="c3a7b-764">必要条件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-764">Requirement</span></span>|<span data-ttu-id="c3a7b-765">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-765">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-766">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-766">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-767">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-767">1.0</span></span>|
|[<span data-ttu-id="c3a7b-768">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-768">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-769">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-769">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-770">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-770">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-771">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-771">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-772">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-772">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="c3a7b-773">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-773">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="c3a7b-774">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-774">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a7b-775">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-775">Type</span></span>

*   [<span data-ttu-id="c3a7b-776">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="c3a7b-776">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="c3a7b-777">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-777">Requirements</span></span>

|<span data-ttu-id="c3a7b-778">必要条件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-778">Requirement</span></span>|<span data-ttu-id="c3a7b-779">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-779">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-780">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-780">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-781">1.3</span><span class="sxs-lookup"><span data-stu-id="c3a7b-781">1.3</span></span>|
|[<span data-ttu-id="c3a7b-782">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-782">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-783">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-783">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-784">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-784">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-785">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-785">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-786">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-786">Example</span></span>

```js
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="c3a7b-787">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-787">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="c3a7b-788">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-788">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="c3a7b-789">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-789">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a7b-790">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-790">Read mode</span></span>

<span data-ttu-id="c3a7b-791">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-791">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="c3a7b-792">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-792">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c3a7b-793">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-793">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a7b-794">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-794">Compose mode</span></span>

<span data-ttu-id="c3a7b-795">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-795">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="c3a7b-796">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-796">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c3a7b-797">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-797">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="c3a7b-798">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-798">Get 500 members maximum.</span></span>
- <span data-ttu-id="c3a7b-799">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-799">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c3a7b-800">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-800">Type</span></span>

*   <span data-ttu-id="c3a7b-801">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-801">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-802">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-802">Requirements</span></span>

|<span data-ttu-id="c3a7b-803">必要条件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-803">Requirement</span></span>|<span data-ttu-id="c3a7b-804">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-804">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-805">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-805">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-806">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-806">1.0</span></span>|
|[<span data-ttu-id="c3a7b-807">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-807">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-808">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-808">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-809">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-809">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-810">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-810">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="c3a7b-811">開催者: [emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)|[開催者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-811">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="c3a7b-812">指定した会議の開催者の電子メールアドレスを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-812">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a7b-813">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-813">Read mode</span></span>

<span data-ttu-id="c3a7b-814">プロパティ`organizer`は、会議の開催者を表す[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-814">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a7b-815">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-815">Compose mode</span></span>

<span data-ttu-id="c3a7b-816">プロパティ`organizer`は、開催者の値を取得するためのメソッドを提供する[オーガナイザー](/javascript/api/outlook/office.organizer)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-816">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="c3a7b-817">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-817">Type</span></span>

*   <span data-ttu-id="c3a7b-818">[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails) | [開催者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-818">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-819">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-819">Requirements</span></span>

|<span data-ttu-id="c3a7b-820">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-820">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="c3a7b-821">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-821">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-822">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-822">1.0</span></span>|<span data-ttu-id="c3a7b-823">1.7</span><span class="sxs-lookup"><span data-stu-id="c3a7b-823">1.7</span></span>|
|[<span data-ttu-id="c3a7b-824">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-824">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-825">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-825">ReadItem</span></span>|<span data-ttu-id="c3a7b-826">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-826">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3a7b-827">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-827">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-828">Read</span><span class="sxs-lookup"><span data-stu-id="c3a7b-828">Read</span></span>|<span data-ttu-id="c3a7b-829">Compose</span><span class="sxs-lookup"><span data-stu-id="c3a7b-829">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="c3a7b-830">(nullable) 定期的なスケジュール:[定期的](/javascript/api/outlook/office.recurrence)なアイテム</span><span class="sxs-lookup"><span data-stu-id="c3a7b-830">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="c3a7b-831">予定の定期的なパターンを取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-831">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="c3a7b-832">会議出席依頼の定期的なパターンを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-832">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="c3a7b-833">予定アイテムの読み取りおよび作成モード。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-833">Read and compose modes for appointment items.</span></span> <span data-ttu-id="c3a7b-834">会議出席依頼アイテムの閲覧モード。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-834">Read mode for meeting request items.</span></span>

<span data-ttu-id="c3a7b-835">この`recurrence`プロパティは、アイテムが series または series 内のインスタンスの場合、定期的な予定または会議出席依頼に対して[定期的](/javascript/api/outlook/office.recurrence)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-835">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="c3a7b-836">`null`は、単一の予定および1つの予定の会議出席依頼に対して返されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-836">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="c3a7b-837">`undefined`は、会議出席依頼ではないメッセージに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-837">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="c3a7b-838">注: 会議出席依頼に`itemClass`は、IPM という値があります。出席依頼。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-838">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="c3a7b-839">注: 定期的なオブジェクトが`null`の場合は、そのオブジェクトが単一の予定または1つの予定の会議出席依頼であり、データ系列の一部ではないことを示します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-839">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a7b-840">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-840">Read mode</span></span>

<span data-ttu-id="c3a7b-841">この`recurrence`プロパティは、定期的な予定を表す[定期的](/javascript/api/outlook/office.recurrence)オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-841">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="c3a7b-842">これは、予定および会議出席依頼に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-842">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a7b-843">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-843">Compose mode</span></span>

<span data-ttu-id="c3a7b-844">この`recurrence`プロパティは、予定の繰り返しを管理するためのメソッドを提供する[定期的](/javascript/api/outlook/office.recurrence)なオブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-844">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="c3a7b-845">これは予定に対して使用できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-845">This is available for appointments.</span></span>

```js
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

##### <a name="type"></a><span data-ttu-id="c3a7b-846">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-846">Type</span></span>

* [<span data-ttu-id="c3a7b-847">繰り返さ</span><span class="sxs-lookup"><span data-stu-id="c3a7b-847">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="c3a7b-848">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-848">Requirement</span></span>|<span data-ttu-id="c3a7b-849">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-849">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-850">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-850">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-851">1.7</span><span class="sxs-lookup"><span data-stu-id="c3a7b-851">1.7</span></span>|
|[<span data-ttu-id="c3a7b-852">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-852">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-853">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-853">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-854">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-854">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-855">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-855">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="c3a7b-856">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-856">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="c3a7b-857">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-857">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="c3a7b-858">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-858">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a7b-859">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-859">Read mode</span></span>

<span data-ttu-id="c3a7b-860">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-860">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="c3a7b-861">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-861">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c3a7b-862">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-862">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a7b-863">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-863">Compose mode</span></span>

<span data-ttu-id="c3a7b-864">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-864">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="c3a7b-865">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-865">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c3a7b-866">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-866">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="c3a7b-867">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-867">Get 500 members maximum.</span></span>
- <span data-ttu-id="c3a7b-868">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-868">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="c3a7b-869">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-869">Type</span></span>

*   <span data-ttu-id="c3a7b-870">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-870">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-871">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-871">Requirements</span></span>

|<span data-ttu-id="c3a7b-872">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-872">Requirement</span></span>|<span data-ttu-id="c3a7b-873">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-873">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-874">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-874">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-875">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-875">1.0</span></span>|
|[<span data-ttu-id="c3a7b-876">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-876">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-877">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-877">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-878">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-878">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-879">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-879">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="c3a7b-880">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-880">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="c3a7b-p135">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p135">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="c3a7b-p136">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsfrom) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p136">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a7b-885">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-885">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a7b-886">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-886">Type</span></span>

*   [<span data-ttu-id="c3a7b-887">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c3a7b-887">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="c3a7b-888">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-888">Requirements</span></span>

|<span data-ttu-id="c3a7b-889">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-889">Requirement</span></span>|<span data-ttu-id="c3a7b-890">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-890">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-891">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-891">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-892">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-892">1.0</span></span>|
|[<span data-ttu-id="c3a7b-893">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-893">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-894">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-894">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-895">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-895">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-896">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-896">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-897">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-897">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="c3a7b-898">(nullable) 系列 Id: String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-898">(nullable) seriesId: String</span></span>

<span data-ttu-id="c3a7b-899">インスタンスが属する系列の id を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-899">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="c3a7b-900">Web 上の Outlook およびデスクトップクライアントでは、 `seriesId`は、このアイテムが属する親 (シリーズ) アイテムの Exchange web サービス (EWS) ID を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-900">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="c3a7b-901">ただし、iOS と Android では、 `seriesId`は親アイテムの REST ID を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-901">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a7b-902">`seriesId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-902">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="c3a7b-903">`seriesId`プロパティが OUTLOOK REST API で使用される outlook id と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-903">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="c3a7b-904">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-904">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="c3a7b-905">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-905">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="c3a7b-906">この`seriesId`プロパティは`null` 、単一の予定、系列のアイテム、会議出席依頼などの親アイテムを持たないアイテムに`undefined`対して、会議出席依頼以外のアイテムに対して返されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-906">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="c3a7b-907">Type</span><span class="sxs-lookup"><span data-stu-id="c3a7b-907">Type</span></span>

* <span data-ttu-id="c3a7b-908">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-908">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-909">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-909">Requirements</span></span>

|<span data-ttu-id="c3a7b-910">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-910">Requirement</span></span>|<span data-ttu-id="c3a7b-911">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-911">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-912">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-912">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-913">1.7</span><span class="sxs-lookup"><span data-stu-id="c3a7b-913">1.7</span></span>|
|[<span data-ttu-id="c3a7b-914">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-914">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-915">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-915">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-916">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-916">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-917">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-917">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-918">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-918">Example</span></span>

```js
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="c3a7b-919">start: Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-919">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="c3a7b-920">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-920">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="c3a7b-p139">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p139">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a7b-923">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-923">Read mode</span></span>

<span data-ttu-id="c3a7b-924">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-924">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a7b-925">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-925">Compose mode</span></span>

<span data-ttu-id="c3a7b-926">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-926">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="c3a7b-927">[`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-927">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="c3a7b-928">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-928">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="c3a7b-929">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-929">Type</span></span>

*   <span data-ttu-id="c3a7b-930">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-930">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-931">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-931">Requirements</span></span>

|<span data-ttu-id="c3a7b-932">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-932">Requirement</span></span>|<span data-ttu-id="c3a7b-933">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-933">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-934">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-934">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-935">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-935">1.0</span></span>|
|[<span data-ttu-id="c3a7b-936">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-936">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-937">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-937">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-938">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-938">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-939">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-939">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="c3a7b-940">subject: String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-940">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="c3a7b-941">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-941">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="c3a7b-942">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-942">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a7b-943">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-943">Read mode</span></span>

<span data-ttu-id="c3a7b-p140">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p140">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="c3a7b-946">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-946">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a7b-947">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-947">Compose mode</span></span>
<span data-ttu-id="c3a7b-948">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-948">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="c3a7b-949">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-949">Type</span></span>

*   <span data-ttu-id="c3a7b-950">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-950">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-951">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-951">Requirements</span></span>

|<span data-ttu-id="c3a7b-952">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-952">Requirement</span></span>|<span data-ttu-id="c3a7b-953">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-953">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-954">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-954">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-955">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-955">1.0</span></span>|
|[<span data-ttu-id="c3a7b-956">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-956">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-957">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-957">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-958">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-958">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-959">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-959">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="c3a7b-960">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-960">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="c3a7b-961">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-961">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="c3a7b-962">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-962">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c3a7b-963">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-963">Read mode</span></span>

<span data-ttu-id="c3a7b-964">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-964">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="c3a7b-965">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-965">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c3a7b-966">ただし、Windows および Mac では、最大 500 人のメンバーを取得できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-966">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="c3a7b-967">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-967">Compose mode</span></span>

<span data-ttu-id="c3a7b-968">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-968">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="c3a7b-969">既定では、コレクションは最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-969">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="c3a7b-970">ただし、Windows および Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-970">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="c3a7b-971">最大 500 人のメンバーを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-971">Get 500 members maximum.</span></span>
- <span data-ttu-id="c3a7b-972">呼び出しごとに最大 100 人のメンバーを設定し、合計で最大 500 人のメンバーを設定します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-972">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c3a7b-973">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-973">Type</span></span>

*   <span data-ttu-id="c3a7b-974">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-974">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-975">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-975">Requirements</span></span>

|<span data-ttu-id="c3a7b-976">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-976">Requirement</span></span>|<span data-ttu-id="c3a7b-977">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-977">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-978">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-978">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-979">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-979">1.0</span></span>|
|[<span data-ttu-id="c3a7b-980">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-980">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-981">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-981">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-982">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-982">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-983">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-983">Compose or Read</span></span>|

## <a name="method-details"></a><span data-ttu-id="c3a7b-984">メソッドの詳細</span><span class="sxs-lookup"><span data-stu-id="c3a7b-984">Method details</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="c3a7b-985">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c3a7b-985">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c3a7b-986">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-986">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="c3a7b-987">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-987">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="c3a7b-988">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-988">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a7b-989">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a7b-989">Parameters</span></span>
|<span data-ttu-id="c3a7b-990">名前</span><span class="sxs-lookup"><span data-stu-id="c3a7b-990">Name</span></span>|<span data-ttu-id="c3a7b-991">種類</span><span class="sxs-lookup"><span data-stu-id="c3a7b-991">Type</span></span>|<span data-ttu-id="c3a7b-992">属性</span><span class="sxs-lookup"><span data-stu-id="c3a7b-992">Attributes</span></span>|<span data-ttu-id="c3a7b-993">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-993">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="c3a7b-994">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-994">String</span></span>||<span data-ttu-id="c3a7b-p144">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p144">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="c3a7b-997">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-997">String</span></span>||<span data-ttu-id="c3a7b-p145">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p145">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="c3a7b-1000">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1000">Object</span></span>|<span data-ttu-id="c3a7b-1001">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1001">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1002">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1002">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a7b-1003">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1003">Object</span></span>|<span data-ttu-id="c3a7b-1004">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1005">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1005">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="c3a7b-1006">Boolean</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1006">Boolean</span></span>|<span data-ttu-id="c3a7b-1007">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1008">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1008">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="c3a7b-1009">function</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1009">function</span></span>|<span data-ttu-id="c3a7b-1010">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1010">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1011">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1011">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c3a7b-1012">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1012">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c3a7b-1013">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1013">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c3a7b-1014">エラー</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1014">Errors</span></span>

|<span data-ttu-id="c3a7b-1015">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1015">Error code</span></span>|<span data-ttu-id="c3a7b-1016">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1016">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="c3a7b-1017">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1017">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="c3a7b-1018">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1018">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="c3a7b-1019">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1019">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1020">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1020">Requirements</span></span>

|<span data-ttu-id="c3a7b-1021">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1021">Requirement</span></span>|<span data-ttu-id="c3a7b-1022">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1022">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1023">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1023">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1024">1.1</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1024">1.1</span></span>|
|[<span data-ttu-id="c3a7b-1025">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1025">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1026">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1026">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3a7b-1027">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1027">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1028">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1028">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c3a7b-1029">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1029">Examples</span></span>

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

<span data-ttu-id="c3a7b-1030">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1030">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
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

<br>

---
---

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="c3a7b-1031">addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1031">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c3a7b-1032">Base64 エンコードのファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1032">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="c3a7b-1033">この`addFileAttachmentFromBase64Async`メソッドは、base64 エンコードからファイルをアップロードし、新規作成フォームのアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1033">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="c3a7b-1034">このメソッドは、AsyncResult オブジェクトの添付ファイル識別子を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1034">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="c3a7b-1035">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1035">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a7b-1036">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1036">Parameters</span></span>

|<span data-ttu-id="c3a7b-1037">名前</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1037">Name</span></span>|<span data-ttu-id="c3a7b-1038">種類</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1038">Type</span></span>|<span data-ttu-id="c3a7b-1039">属性</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1039">Attributes</span></span>|<span data-ttu-id="c3a7b-1040">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1040">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="c3a7b-1041">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1041">String</span></span>||<span data-ttu-id="c3a7b-1042">電子メールまたはイベントに追加する画像またはファイルの、base64 でエンコードされたコンテンツ。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1042">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="c3a7b-1043">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1043">String</span></span>||<span data-ttu-id="c3a7b-p147">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p147">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="c3a7b-1046">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1046">Object</span></span>|<span data-ttu-id="c3a7b-1047">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1047">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1048">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1048">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a7b-1049">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1049">Object</span></span>|<span data-ttu-id="c3a7b-1050">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1050">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1051">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1051">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="c3a7b-1052">Boolean</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1052">Boolean</span></span>|<span data-ttu-id="c3a7b-1053">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1053">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1054">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1054">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="c3a7b-1055">function</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1055">function</span></span>|<span data-ttu-id="c3a7b-1056">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1056">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1057">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1057">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c3a7b-1058">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1058">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c3a7b-1059">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1059">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c3a7b-1060">エラー</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1060">Errors</span></span>

|<span data-ttu-id="c3a7b-1061">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1061">Error code</span></span>|<span data-ttu-id="c3a7b-1062">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1062">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="c3a7b-1063">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1063">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="c3a7b-1064">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1064">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="c3a7b-1065">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1065">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1066">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1066">Requirements</span></span>

|<span data-ttu-id="c3a7b-1067">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1067">Requirement</span></span>|<span data-ttu-id="c3a7b-1068">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1068">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1069">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1069">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1070">1.8</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1070">1.8</span></span>|
|[<span data-ttu-id="c3a7b-1071">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1071">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1072">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1072">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3a7b-1073">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1073">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1074">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1074">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c3a7b-1075">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1075">Examples</span></span>

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
        // Do something here.
      });
  });
```

<br>

---
---

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="c3a7b-1076">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1076">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="c3a7b-1077">サポートされているイベントのイベント ハンドラーを追加します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1077">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="c3a7b-1078">現在、サポートされて`Office.EventType.AttachmentsChanged`いる`Office.EventType.AppointmentTimeChanged`イベント`Office.EventType.EnhancedLocationsChanged`の`Office.EventType.RecipientsChanged`種類は`Office.EventType.RecurrenceChanged`、、、、、です。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1078">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a7b-1079">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1079">Parameters</span></span>

| <span data-ttu-id="c3a7b-1080">名前</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1080">Name</span></span> | <span data-ttu-id="c3a7b-1081">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1081">Type</span></span> | <span data-ttu-id="c3a7b-1082">属性</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1082">Attributes</span></span> | <span data-ttu-id="c3a7b-1083">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1083">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c3a7b-1084">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1084">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c3a7b-1085">ハンドラーを呼び出す必要のあるイベント。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1085">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="c3a7b-1086">Function</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1086">Function</span></span> || <span data-ttu-id="c3a7b-p148">イベントを処理する関数。関数は、オブジェクト リテラルである単一パラメーターを受け入れる必要があります。パラメーターの `type` プロパティは、`addHandlerAsync` に渡される `eventType` パラメーターと一致します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p148">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="c3a7b-1090">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1090">Object</span></span> | <span data-ttu-id="c3a7b-1091">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1091">&lt;optional&gt;</span></span> | <span data-ttu-id="c3a7b-1092">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1092">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c3a7b-1093">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1093">Object</span></span> | <span data-ttu-id="c3a7b-1094">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1094">&lt;optional&gt;</span></span> | <span data-ttu-id="c3a7b-1095">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1095">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c3a7b-1096">function</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1096">function</span></span>| <span data-ttu-id="c3a7b-1097">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1098">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1098">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1099">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1099">Requirements</span></span>

|<span data-ttu-id="c3a7b-1100">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1100">Requirement</span></span>| <span data-ttu-id="c3a7b-1101">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1101">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1102">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1102">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3a7b-1103">1.7</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1103">1.7</span></span> |
|[<span data-ttu-id="c3a7b-1104">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1104">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3a7b-1105">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1105">ReadItem</span></span> |
|[<span data-ttu-id="c3a7b-1106">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1106">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c3a7b-1107">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1107">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="c3a7b-1108">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1108">Example</span></span>

```js
function myHandlerFunction(eventarg) {
  if (eventarg.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added) {
    var attachment = eventarg.attachmentDetails;
    console.log("Event Fired and Attachment Added!");
    getAttachmentContentAsync(attachment.id, options, callback);
  }
}

Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, myHandlerFunction, myCallback);
```

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="c3a7b-1109">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1109">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c3a7b-1110">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1110">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="c3a7b-p149">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p149">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="c3a7b-1114">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1114">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="c3a7b-1115">Office アドインを Outlook on the web で実行している場合、編集中のアイテム以外のアイテムに `addItemAttachmentAsync` メソッドでアイテムを添付できます。ただし、これはサポートされていないため、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1115">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a7b-1116">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1116">Parameters</span></span>

|<span data-ttu-id="c3a7b-1117">名前</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1117">Name</span></span>|<span data-ttu-id="c3a7b-1118">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1118">Type</span></span>|<span data-ttu-id="c3a7b-1119">属性</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1119">Attributes</span></span>|<span data-ttu-id="c3a7b-1120">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1120">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="c3a7b-1121">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1121">String</span></span>||<span data-ttu-id="c3a7b-p150">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p150">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="c3a7b-1124">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1124">String</span></span>||<span data-ttu-id="c3a7b-1125">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1125">The subject of the item to be attached.</span></span> <span data-ttu-id="c3a7b-1126">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1126">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="c3a7b-1127">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1127">Object</span></span>|<span data-ttu-id="c3a7b-1128">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1128">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1129">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1129">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a7b-1130">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1130">Object</span></span>|<span data-ttu-id="c3a7b-1131">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1131">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1132">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1132">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c3a7b-1133">function</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1133">function</span></span>|<span data-ttu-id="c3a7b-1134">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1134">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1135">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1135">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c3a7b-1136">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1136">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c3a7b-1137">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1137">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c3a7b-1138">エラー</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1138">Errors</span></span>

|<span data-ttu-id="c3a7b-1139">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1139">Error code</span></span>|<span data-ttu-id="c3a7b-1140">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1140">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="c3a7b-1141">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1141">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1142">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1142">Requirements</span></span>

|<span data-ttu-id="c3a7b-1143">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1143">Requirement</span></span>|<span data-ttu-id="c3a7b-1144">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1144">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1145">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1146">1.1</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1146">1.1</span></span>|
|[<span data-ttu-id="c3a7b-1147">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1147">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1148">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1148">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3a7b-1149">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1149">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1150">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-1151">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1151">Example</span></span>

<span data-ttu-id="c3a7b-1152">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1152">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="c3a7b-1153">close()</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1153">close()</span></span>

<span data-ttu-id="c3a7b-1154">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1154">Closes the current item that is being composed.</span></span>

<span data-ttu-id="c3a7b-p152">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p152">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a7b-1157">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1157">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="c3a7b-1158">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1158">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1159">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1159">Requirements</span></span>

|<span data-ttu-id="c3a7b-1160">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1160">Requirement</span></span>|<span data-ttu-id="c3a7b-1161">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1161">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1162">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1162">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1163">1.3</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1163">1.3</span></span>|
|[<span data-ttu-id="c3a7b-1164">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1164">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1165">制限あり</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1165">Restricted</span></span>|
|[<span data-ttu-id="c3a7b-1166">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1166">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1167">新規作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1167">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="c3a7b-1168">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1168">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="c3a7b-1169">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1169">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a7b-1170">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1170">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c3a7b-1171">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1171">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c3a7b-1172">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1172">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="c3a7b-p153">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a7b-1176">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1176">Parameters</span></span>

|<span data-ttu-id="c3a7b-1177">名前</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1177">Name</span></span>|<span data-ttu-id="c3a7b-1178">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1178">Type</span></span>|<span data-ttu-id="c3a7b-1179">属性</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1179">Attributes</span></span>|<span data-ttu-id="c3a7b-1180">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1180">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="c3a7b-1181">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1181">String &#124; Object</span></span>||<span data-ttu-id="c3a7b-p154">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c3a7b-1184">**または**</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1184">**OR**</span></span><br/><span data-ttu-id="c3a7b-p155">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="c3a7b-1187">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1187">String</span></span>|<span data-ttu-id="c3a7b-1188">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1188">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-p156">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="c3a7b-1191">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1191">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="c3a7b-1192">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1192">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1193">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1193">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="c3a7b-1194">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1194">String</span></span>||<span data-ttu-id="c3a7b-p157">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="c3a7b-1197">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1197">String</span></span>||<span data-ttu-id="c3a7b-1198">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1198">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="c3a7b-1199">文字列</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1199">String</span></span>||<span data-ttu-id="c3a7b-p158">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="c3a7b-1202">ブール値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1202">Boolean</span></span>||<span data-ttu-id="c3a7b-p159">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="c3a7b-1205">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1205">String</span></span>||<span data-ttu-id="c3a7b-p160">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="c3a7b-1209">function</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1209">function</span></span>|<span data-ttu-id="c3a7b-1210">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1210">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1211">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1211">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1212">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1212">Requirements</span></span>

|<span data-ttu-id="c3a7b-1213">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1213">Requirement</span></span>|<span data-ttu-id="c3a7b-1214">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1214">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1215">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1215">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1216">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1216">1.0</span></span>|
|[<span data-ttu-id="c3a7b-1217">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1217">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1218">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1218">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-1219">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1220">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1220">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c3a7b-1221">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1221">Examples</span></span>

<span data-ttu-id="c3a7b-1222">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1222">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="c3a7b-1223">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1223">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="c3a7b-1224">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1224">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c3a7b-1225">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1225">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c3a7b-1226">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1226">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c3a7b-1227">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1227">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="c3a7b-1228">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1228">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="c3a7b-1229">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1229">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a7b-1230">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1230">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c3a7b-1231">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1231">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c3a7b-1232">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1232">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="c3a7b-p161">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p161">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a7b-1236">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1236">Parameters</span></span>

|<span data-ttu-id="c3a7b-1237">名前</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1237">Name</span></span>|<span data-ttu-id="c3a7b-1238">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1238">Type</span></span>|<span data-ttu-id="c3a7b-1239">属性</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1239">Attributes</span></span>|<span data-ttu-id="c3a7b-1240">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1240">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="c3a7b-1241">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1241">String &#124; Object</span></span>||<span data-ttu-id="c3a7b-p162">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p162">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c3a7b-1244">**または**</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1244">**OR**</span></span><br/><span data-ttu-id="c3a7b-p163">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p163">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="c3a7b-1247">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1247">String</span></span>|<span data-ttu-id="c3a7b-1248">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1248">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-p164">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p164">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="c3a7b-1251">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1251">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="c3a7b-1252">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1252">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1253">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1253">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="c3a7b-1254">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1254">String</span></span>||<span data-ttu-id="c3a7b-p165">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p165">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="c3a7b-1257">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1257">String</span></span>||<span data-ttu-id="c3a7b-1258">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1258">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="c3a7b-1259">文字列</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1259">String</span></span>||<span data-ttu-id="c3a7b-p166">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p166">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="c3a7b-1262">ブール値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1262">Boolean</span></span>||<span data-ttu-id="c3a7b-p167">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p167">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="c3a7b-1265">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1265">String</span></span>||<span data-ttu-id="c3a7b-p168">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p168">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="c3a7b-1269">function</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1269">function</span></span>|<span data-ttu-id="c3a7b-1270">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1270">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1271">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1271">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1272">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1272">Requirements</span></span>

|<span data-ttu-id="c3a7b-1273">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1273">Requirement</span></span>|<span data-ttu-id="c3a7b-1274">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1274">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1275">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1275">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1276">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1276">1.0</span></span>|
|[<span data-ttu-id="c3a7b-1277">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1277">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1278">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1278">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-1279">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1279">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1280">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1280">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c3a7b-1281">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1281">Examples</span></span>

<span data-ttu-id="c3a7b-1282">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1282">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="c3a7b-1283">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1283">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="c3a7b-1284">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1284">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c3a7b-1285">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1285">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c3a7b-1286">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1286">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c3a7b-1287">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1287">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getallinternetheadersasyncoptions-callback"></a><span data-ttu-id="c3a7b-1288">getAllInternetHeadersAsync ([オプション], [callback])</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1288">getAllInternetHeadersAsync([options], [callback])</span></span>

<span data-ttu-id="c3a7b-1289">メッセージのすべてのインターネットヘッダーを文字列として取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1289">Gets all the internet headers for the message as a string.</span></span> <span data-ttu-id="c3a7b-1290">閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1290">Read mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a7b-1291">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1291">Parameters</span></span>

|<span data-ttu-id="c3a7b-1292">名前</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1292">Name</span></span>|<span data-ttu-id="c3a7b-1293">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1293">Type</span></span>|<span data-ttu-id="c3a7b-1294">属性</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1294">Attributes</span></span>|<span data-ttu-id="c3a7b-1295">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1295">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c3a7b-1296">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1296">Object</span></span>|<span data-ttu-id="c3a7b-1297">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1297">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1298">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1298">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a7b-1299">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1299">Object</span></span>|<span data-ttu-id="c3a7b-1300">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1300">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1301">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1301">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c3a7b-1302">function</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1302">function</span></span>|<span data-ttu-id="c3a7b-1303">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1303">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1304">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1304">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> <span data-ttu-id="c3a7b-1305">成功した場合、インターネットヘッダーデータは、文字列として asyncResult プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1305">On success, the internet headers data is provided in the asyncResult.value property as a string.</span></span> <span data-ttu-id="c3a7b-1306">返される文字列値の書式情報については、 [RFC 2183](https://tools.ietf.org/html/rfc2183)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1306">Refer to [RFC 2183](https://tools.ietf.org/html/rfc2183) for the formatting information of the returned string value.</span></span> <span data-ttu-id="c3a7b-1307">呼び出しが失敗した場合、asyncResult. error プロパティには、エラーの理由と共にエラーコードが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1307">If the call fails, the asyncResult.error property will contain an error code with the reason for the failure.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1308">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1308">Requirements</span></span>

|<span data-ttu-id="c3a7b-1309">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1309">Requirement</span></span>|<span data-ttu-id="c3a7b-1310">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1310">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1312">1.8</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1312">1.8</span></span>|
|[<span data-ttu-id="c3a7b-1313">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1314">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-1315">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1316">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1316">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3a7b-1317">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1317">Returns:</span></span>

<span data-ttu-id="c3a7b-1318">[RFC 2183](https://tools.ietf.org/html/rfc2183)に従って書式設定された文字列としてのインターネットヘッダーデータ。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1318">The internet headers data as a string formatted according to [RFC 2183](https://tools.ietf.org/html/rfc2183).</span></span>

<span data-ttu-id="c3a7b-1319">型:String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1319">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="c3a7b-1320">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1320">Example</span></span>

```js
// Get the internet headers related to the mail.
Office.context.mailbox.item.getAllInternetHeadersAsync(
  function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log(asyncResult.value);
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

<br>

---
---

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="c3a7b-1321">getAttachmentContentAsync (attachmentId, [options], [callback]) > [Attachmentcontent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1321">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="c3a7b-1322">メッセージまたは予定から指定された添付ファイルを取得し`AttachmentContent` 、それをオブジェクトとして返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1322">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="c3a7b-1323">メソッド`getAttachmentContentAsync`は、指定された id の添付ファイルをアイテムから取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1323">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="c3a7b-1324">ベストプラクティスとして、識別子を使用して、または`getAttachmentsAsync` `item.attachments`の呼び出しで attachmentIds を取得したのと同じセッションの添付ファイルを取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1324">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="c3a7b-1325">Outlook on the web とモバイル デバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1325">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="c3a7b-1326">ユーザーがアプリを閉じたとき、またはインラインフォームの作成が開始されたときに、別のウィンドウで続行するためにフォームをポップアウトした後、セッションが終了します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1326">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a7b-1327">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1327">Parameters</span></span>

|<span data-ttu-id="c3a7b-1328">名前</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1328">Name</span></span>|<span data-ttu-id="c3a7b-1329">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1329">Type</span></span>|<span data-ttu-id="c3a7b-1330">属性</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1330">Attributes</span></span>|<span data-ttu-id="c3a7b-1331">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1331">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="c3a7b-1332">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1332">String</span></span>||<span data-ttu-id="c3a7b-1333">取得する添付ファイルの識別子を指定します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1333">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="c3a7b-1334">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1334">Object</span></span>|<span data-ttu-id="c3a7b-1335">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1335">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1336">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1336">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a7b-1337">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1337">Object</span></span>|<span data-ttu-id="c3a7b-1338">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1338">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1339">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1339">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c3a7b-1340">function</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1340">function</span></span>|<span data-ttu-id="c3a7b-1341">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1341">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1342">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1342">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1343">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1343">Requirements</span></span>

|<span data-ttu-id="c3a7b-1344">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1344">Requirement</span></span>|<span data-ttu-id="c3a7b-1345">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1345">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1346">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1347">1.8</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1347">1.8</span></span>|
|[<span data-ttu-id="c3a7b-1348">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1349">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-1350">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1351">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1351">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3a7b-1352">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1352">Returns:</span></span>

<span data-ttu-id="c3a7b-1353">型: [Attachmentcontent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1353">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="c3a7b-1354">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1354">Example</span></span>

```js
var item = Office.context.mailbox.item;
var listOfAttachments = [];
var options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      // Handle file attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Handle email item attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Handle .icalender attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      break;
    default:
      // Handle attachment formats that are not supported.
  }
}
```

<br>

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="c3a7b-1355">getAttachmentsAsync ([オプション], [callback]) > Array. <[Attachmentdetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c3a7b-1355">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="c3a7b-1356">アイテムの添付ファイルを配列として取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1356">Gets the item's attachments as an array.</span></span> <span data-ttu-id="c3a7b-1357">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1357">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a7b-1358">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1358">Parameters</span></span>

|<span data-ttu-id="c3a7b-1359">名前</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1359">Name</span></span>|<span data-ttu-id="c3a7b-1360">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1360">Type</span></span>|<span data-ttu-id="c3a7b-1361">属性</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1361">Attributes</span></span>|<span data-ttu-id="c3a7b-1362">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1362">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c3a7b-1363">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1363">Object</span></span>|<span data-ttu-id="c3a7b-1364">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1364">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1365">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1365">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a7b-1366">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1366">Object</span></span>|<span data-ttu-id="c3a7b-1367">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1367">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1368">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1368">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c3a7b-1369">function</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1369">function</span></span>|<span data-ttu-id="c3a7b-1370">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1370">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1371">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1371">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1372">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1372">Requirements</span></span>

|<span data-ttu-id="c3a7b-1373">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1373">Requirement</span></span>|<span data-ttu-id="c3a7b-1374">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1374">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1375">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1376">1.8</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1376">1.8</span></span>|
|[<span data-ttu-id="c3a7b-1377">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1377">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1378">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-1379">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1379">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1380">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1380">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3a7b-1381">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1381">Returns:</span></span>

<span data-ttu-id="c3a7b-1382">型: Array. <[attachmentdetails 詳細](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c3a7b-1382">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="c3a7b-1383">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1383">Example</span></span>

<span data-ttu-id="c3a7b-1384">次の例では、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1384">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```js
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

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="c3a7b-1385">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1385">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="c3a7b-1386">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1386">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a7b-1387">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1387">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1388">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1388">Requirements</span></span>

|<span data-ttu-id="c3a7b-1389">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1389">Requirement</span></span>|<span data-ttu-id="c3a7b-1390">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1390">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1391">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1391">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1392">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1392">1.0</span></span>|
|[<span data-ttu-id="c3a7b-1393">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1393">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1394">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1394">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-1395">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1395">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1396">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1396">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3a7b-1397">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1397">Returns:</span></span>

<span data-ttu-id="c3a7b-1398">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1398">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="c3a7b-1399">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1399">Example</span></span>

<span data-ttu-id="c3a7b-1400">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1400">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="c3a7b-1401">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1401">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c3a7b-1402">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1402">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a7b-1403">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1403">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a7b-1404">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1404">Parameters</span></span>

|<span data-ttu-id="c3a7b-1405">名前</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1405">Name</span></span>|<span data-ttu-id="c3a7b-1406">種類</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1406">Type</span></span>|<span data-ttu-id="c3a7b-1407">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1407">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="c3a7b-1408">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1408">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="c3a7b-1409">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1409">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1410">Requirements</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1410">Requirements</span></span>

|<span data-ttu-id="c3a7b-1411">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1411">Requirement</span></span>|<span data-ttu-id="c3a7b-1412">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1412">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1413">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1414">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1414">1.0</span></span>|
|[<span data-ttu-id="c3a7b-1415">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1416">制限あり</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1416">Restricted</span></span>|
|[<span data-ttu-id="c3a7b-1417">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1418">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1418">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3a7b-1419">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1419">Returns:</span></span>

<span data-ttu-id="c3a7b-1420">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1420">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="c3a7b-1421">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1421">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="c3a7b-1422">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1422">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="c3a7b-1423">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1423">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="c3a7b-1424">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1424">Value of `entityType`</span></span>|<span data-ttu-id="c3a7b-1425">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1425">Type of objects in returned array</span></span>|<span data-ttu-id="c3a7b-1426">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1426">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="c3a7b-1427">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1427">String</span></span>|<span data-ttu-id="c3a7b-1428">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1428">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="c3a7b-1429">連絡先</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1429">Contact</span></span>|<span data-ttu-id="c3a7b-1430">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1430">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="c3a7b-1431">文字列</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1431">String</span></span>|<span data-ttu-id="c3a7b-1432">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1432">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="c3a7b-1433">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1433">MeetingSuggestion</span></span>|<span data-ttu-id="c3a7b-1434">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1434">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="c3a7b-1435">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1435">PhoneNumber</span></span>|<span data-ttu-id="c3a7b-1436">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1436">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="c3a7b-1437">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1437">TaskSuggestion</span></span>|<span data-ttu-id="c3a7b-1438">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1438">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="c3a7b-1439">文字列</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1439">String</span></span>|<span data-ttu-id="c3a7b-1440">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1440">**Restricted**</span></span>|

<span data-ttu-id="c3a7b-1441">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c3a7b-1441">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="c3a7b-1442">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1442">Example</span></span>

<span data-ttu-id="c3a7b-1443">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1443">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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
};
```

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="c3a7b-1444">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1444">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c3a7b-1445">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1445">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a7b-1446">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1446">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c3a7b-1447">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1447">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a7b-1448">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1448">Parameters</span></span>

|<span data-ttu-id="c3a7b-1449">名前</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1449">Name</span></span>|<span data-ttu-id="c3a7b-1450">種類</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1450">Type</span></span>|<span data-ttu-id="c3a7b-1451">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1451">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="c3a7b-1452">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1452">String</span></span>|<span data-ttu-id="c3a7b-1453">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1453">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1454">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1454">Requirements</span></span>

|<span data-ttu-id="c3a7b-1455">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1455">Requirement</span></span>|<span data-ttu-id="c3a7b-1456">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1456">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1457">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1457">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1458">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1458">1.0</span></span>|
|[<span data-ttu-id="c3a7b-1459">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1459">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1460">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1460">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-1461">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1461">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1462">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1462">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3a7b-1463">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1463">Returns:</span></span>

<span data-ttu-id="c3a7b-p174">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p174">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="c3a7b-1466">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c3a7b-1466">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

<br>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="c3a7b-1467">、Office.context.mailbox.item.getinitializationcontextasync ([オプション], [callback])</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1467">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="c3a7b-1468">[アクション可能なメッセージによってアドインがアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されたときに渡される初期化データを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1468">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="c3a7b-1469">このメソッドは、Outlook 2016 以降の Windows (16.0.8413.1000 より後のバージョン) および Outlook on the Office 365 でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1469">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a7b-1470">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1470">Parameters</span></span>

|<span data-ttu-id="c3a7b-1471">名前</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1471">Name</span></span>|<span data-ttu-id="c3a7b-1472">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1472">Type</span></span>|<span data-ttu-id="c3a7b-1473">属性</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1473">Attributes</span></span>|<span data-ttu-id="c3a7b-1474">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1474">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c3a7b-1475">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1475">Object</span></span>|<span data-ttu-id="c3a7b-1476">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1476">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1477">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1477">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a7b-1478">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1478">Object</span></span>|<span data-ttu-id="c3a7b-1479">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1479">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1480">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1480">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c3a7b-1481">function</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1481">function</span></span>|<span data-ttu-id="c3a7b-1482">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1482">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1483">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1483">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c3a7b-1484">成功すると、初期化データが文字列とし`asyncResult.value`てプロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1484">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="c3a7b-1485">初期化コンテキストがない場合、 `asyncResult`オブジェクトには、 `Error` `code`プロパティがに`9020`設定されたオブジェクトと`name`プロパティがに`GenericResponseError`設定されたオブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1485">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1486">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1486">Requirements</span></span>

|<span data-ttu-id="c3a7b-1487">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1487">Requirement</span></span>|<span data-ttu-id="c3a7b-1488">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1488">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1489">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1490">プレビュー</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1490">Preview</span></span>|
|[<span data-ttu-id="c3a7b-1491">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1491">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1492">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1492">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-1493">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1493">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1494">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1494">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-1495">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1495">Example</span></span>

```js
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

<br>

---
---

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="c3a7b-1496">getItemIdAsync ([オプション], callback)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1496">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="c3a7b-1497">保存されたアイテムの ID を非同期に取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1497">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="c3a7b-1498">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1498">Compose mode only.</span></span>

<span data-ttu-id="c3a7b-1499">このメソッドを呼び出すと、コールバックメソッドによってアイテム ID が返されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1499">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a7b-1500">アドインが新規作成モードの`getItemIdAsync`アイテムに対して呼び出しを行う場合 ( `itemId` EWS または REST API を使用するため)、Outlook がキャッシュモードの場合は、アイテムがサーバーに同期されるまでしばらく時間がかかる場合があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1500">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="c3a7b-1501">アイテムが同期されるまで、 `itemId`は認識されず、を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1501">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a7b-1502">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1502">Parameters</span></span>

|<span data-ttu-id="c3a7b-1503">名前</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1503">Name</span></span>|<span data-ttu-id="c3a7b-1504">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1504">Type</span></span>|<span data-ttu-id="c3a7b-1505">属性</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1505">Attributes</span></span>|<span data-ttu-id="c3a7b-1506">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1506">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c3a7b-1507">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1507">Object</span></span>|<span data-ttu-id="c3a7b-1508">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1508">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1509">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1509">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a7b-1510">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1510">Object</span></span>|<span data-ttu-id="c3a7b-1511">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1511">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1512">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1512">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c3a7b-1513">function</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1513">function</span></span>||<span data-ttu-id="c3a7b-1514">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1514">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c3a7b-1515">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1515">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c3a7b-1516">エラー</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1516">Errors</span></span>

|<span data-ttu-id="c3a7b-1517">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1517">Error code</span></span>|<span data-ttu-id="c3a7b-1518">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1518">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="c3a7b-1519">この id は、アイテムが保存されるまでは取得できません。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1519">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1520">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1520">Requirements</span></span>

|<span data-ttu-id="c3a7b-1521">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1521">Requirement</span></span>|<span data-ttu-id="c3a7b-1522">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1522">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1523">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1523">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1524">1.8</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1524">1.8</span></span>|
|[<span data-ttu-id="c3a7b-1525">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1525">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1526">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1526">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-1527">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1527">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1528">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1528">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c3a7b-1529">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1529">Examples</span></span>

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="c3a7b-1530">次の例は、コールバック関数`result`に渡されるパラメーターの構造を示しています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1530">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="c3a7b-1531">プロパティ`value`には、アイテムの ID が含まれています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1531">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="c3a7b-1532">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1532">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="c3a7b-1533">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1533">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a7b-1534">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1534">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c3a7b-p178">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p178">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c3a7b-1538">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1538">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c3a7b-1539">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1539">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c3a7b-p179">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1543">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1543">Requirements</span></span>

|<span data-ttu-id="c3a7b-1544">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1544">Requirement</span></span>|<span data-ttu-id="c3a7b-1545">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1545">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1546">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1547">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1547">1.0</span></span>|
|[<span data-ttu-id="c3a7b-1548">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1548">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1549">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1549">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-1550">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1550">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1551">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1551">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3a7b-1552">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1552">Returns:</span></span>

<span data-ttu-id="c3a7b-p180">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="c3a7b-1555">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1555">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c3a7b-1556">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1556">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c3a7b-1557">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1557">Example</span></span>

<span data-ttu-id="c3a7b-1558">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1558">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="c3a7b-1559">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1559">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="c3a7b-1560">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1560">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a7b-1561">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1561">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c3a7b-1562">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1562">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="c3a7b-p181">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p181">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a7b-1565">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1565">Parameters</span></span>

|<span data-ttu-id="c3a7b-1566">名前</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1566">Name</span></span>|<span data-ttu-id="c3a7b-1567">種類</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1567">Type</span></span>|<span data-ttu-id="c3a7b-1568">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1568">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="c3a7b-1569">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1569">String</span></span>|<span data-ttu-id="c3a7b-1570">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1570">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1571">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1571">Requirements</span></span>

|<span data-ttu-id="c3a7b-1572">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1572">Requirement</span></span>|<span data-ttu-id="c3a7b-1573">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1573">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1574">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1574">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1575">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1575">1.0</span></span>|
|[<span data-ttu-id="c3a7b-1576">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1576">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1577">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1577">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-1578">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1578">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1579">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1579">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3a7b-1580">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1580">Returns:</span></span>

<span data-ttu-id="c3a7b-1581">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1581">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="c3a7b-1582">型: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="c3a7b-1582">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="c3a7b-1583">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1583">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="c3a7b-1584">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1584">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="c3a7b-1585">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1585">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="c3a7b-p182">選択されていない状態でカーソルが本文または件名にある場合、メソッドは選択されたデータに対し空の文字列を返します。本文または件名以外のフィールドが選択されている場合には、メソッドは`InvalidSelection`エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p182">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a7b-1588">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1588">Parameters</span></span>

|<span data-ttu-id="c3a7b-1589">名前</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1589">Name</span></span>|<span data-ttu-id="c3a7b-1590">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1590">Type</span></span>|<span data-ttu-id="c3a7b-1591">属性</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1591">Attributes</span></span>|<span data-ttu-id="c3a7b-1592">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1592">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="c3a7b-1593">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1593">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="c3a7b-p183">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p183">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="c3a7b-1597">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1597">Object</span></span>|<span data-ttu-id="c3a7b-1598">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1598">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1599">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1599">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a7b-1600">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1600">Object</span></span>|<span data-ttu-id="c3a7b-1601">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1601">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1602">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1602">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c3a7b-1603">function</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1603">function</span></span>||<span data-ttu-id="c3a7b-1604">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1604">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c3a7b-1605">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1605">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="c3a7b-1606">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1606">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1607">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1607">Requirements</span></span>

|<span data-ttu-id="c3a7b-1608">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1608">Requirement</span></span>|<span data-ttu-id="c3a7b-1609">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1609">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1610">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1610">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1611">1.2</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1611">1.2</span></span>|
|[<span data-ttu-id="c3a7b-1612">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1612">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1613">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1613">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-1614">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1614">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1615">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1615">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3a7b-1616">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1616">Returns:</span></span>

<span data-ttu-id="c3a7b-1617">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1617">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="c3a7b-1618">型:String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1618">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="c3a7b-1619">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1619">Example</span></span>

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  console.log("Selected text in " + prop + ": " + text);
}
```

<br>

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="c3a7b-1620">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1620">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="c3a7b-1621">強調表示された一致内で見つかったユーザーが選択しているエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1621">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="c3a7b-1622">強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1622">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="c3a7b-1623">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1623">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1624">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1624">Requirements</span></span>

|<span data-ttu-id="c3a7b-1625">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1625">Requirement</span></span>|<span data-ttu-id="c3a7b-1626">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1626">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1627">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1628">1.6</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1628">1.6</span></span>|
|[<span data-ttu-id="c3a7b-1629">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1629">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1630">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-1631">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1631">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1632">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1632">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3a7b-1633">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1633">Returns:</span></span>

<span data-ttu-id="c3a7b-1634">型:[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1634">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="c3a7b-1635">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1635">Example</span></span>

<span data-ttu-id="c3a7b-1636">次の例では、強調表示された一致内でユーザーが選択した住所エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1636">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="c3a7b-1637">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1637">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="c3a7b-p186">マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返します。強調表示された一致は、[コンテキスト アドイン](/outlook/add-ins/contextual-outlook-add-ins)に適用されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p186">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="c3a7b-1640">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1640">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c3a7b-p187">`getSelectedRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p187">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c3a7b-1644">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1644">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c3a7b-1645">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1645">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c3a7b-p188">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p188">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1649">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1649">Requirements</span></span>

|<span data-ttu-id="c3a7b-1650">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1650">Requirement</span></span>|<span data-ttu-id="c3a7b-1651">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1651">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1652">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1652">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1653">1.6</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1653">1.6</span></span>|
|[<span data-ttu-id="c3a7b-1654">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1654">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1655">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1655">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-1656">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1656">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1657">読み取り</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1657">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c3a7b-1658">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1658">Returns:</span></span>

<span data-ttu-id="c3a7b-p189">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p189">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="c3a7b-1661">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1661">Example</span></span>

<span data-ttu-id="c3a7b-1662">次の例は、マニフェストで指定された正規表現ルールの要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1662">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="c3a7b-1663">getSharedPropertiesAsync ([options], callback)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1663">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="c3a7b-1664">共有フォルダー、予定表、またはメールボックス内の選択した予定またはメッセージのプロパティを取得します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1664">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a7b-1665">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1665">Parameters</span></span>

|<span data-ttu-id="c3a7b-1666">名前</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1666">Name</span></span>|<span data-ttu-id="c3a7b-1667">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1667">Type</span></span>|<span data-ttu-id="c3a7b-1668">属性</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1668">Attributes</span></span>|<span data-ttu-id="c3a7b-1669">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1669">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c3a7b-1670">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1670">Object</span></span>|<span data-ttu-id="c3a7b-1671">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1671">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1672">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1672">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a7b-1673">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1673">Object</span></span>|<span data-ttu-id="c3a7b-1674">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1674">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1675">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1675">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c3a7b-1676">function</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1676">function</span></span>||<span data-ttu-id="c3a7b-1677">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1677">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c3a7b-1678">共有プロパティは、 [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) `asyncResult.value`プロパティのオブジェクトとして提供されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1678">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="c3a7b-1679">このオブジェクトは、アイテムの共有プロパティを取得するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1679">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1680">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1680">Requirements</span></span>

|<span data-ttu-id="c3a7b-1681">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1681">Requirement</span></span>|<span data-ttu-id="c3a7b-1682">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1682">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1683">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1683">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1684">1.8</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1684">1.8</span></span>|
|[<span data-ttu-id="c3a7b-1685">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1685">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1686">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1686">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-1687">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1687">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1688">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1688">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-1689">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1689">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="c3a7b-1690">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1690">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="c3a7b-1691">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1691">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="c3a7b-p191">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p191">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a7b-1695">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1695">Parameters</span></span>

|<span data-ttu-id="c3a7b-1696">名前</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1696">Name</span></span>|<span data-ttu-id="c3a7b-1697">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1697">Type</span></span>|<span data-ttu-id="c3a7b-1698">属性</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1698">Attributes</span></span>|<span data-ttu-id="c3a7b-1699">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1699">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="c3a7b-1700">function</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1700">function</span></span>||<span data-ttu-id="c3a7b-1701">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1701">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c3a7b-1702">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1702">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="c3a7b-1703">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1703">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="c3a7b-1704">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1704">Object</span></span>|<span data-ttu-id="c3a7b-1705">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1705">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1706">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1706">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="c3a7b-1707">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1707">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1708">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1708">Requirements</span></span>

|<span data-ttu-id="c3a7b-1709">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1709">Requirement</span></span>|<span data-ttu-id="c3a7b-1710">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1710">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1711">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1712">1.0</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1712">1.0</span></span>|
|[<span data-ttu-id="c3a7b-1713">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1713">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1714">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1714">ReadItem</span></span>|
|[<span data-ttu-id="c3a7b-1715">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1715">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1716">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1716">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-1717">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1717">Example</span></span>

<span data-ttu-id="c3a7b-p194">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p194">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
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

<br>

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="c3a7b-1721">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1721">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="c3a7b-1722">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1722">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="c3a7b-1723">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1723">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="c3a7b-1724">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1724">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="c3a7b-1725">Outlook on the web とモバイル デバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1725">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="c3a7b-1726">ユーザーがアプリを閉じたとき、またはインラインフォームの作成が開始されたときに、別のウィンドウで続行するためにフォームをポップアウトした後、セッションが終了します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1726">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a7b-1727">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1727">Parameters</span></span>

|<span data-ttu-id="c3a7b-1728">名前</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1728">Name</span></span>|<span data-ttu-id="c3a7b-1729">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1729">Type</span></span>|<span data-ttu-id="c3a7b-1730">属性</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1730">Attributes</span></span>|<span data-ttu-id="c3a7b-1731">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1731">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="c3a7b-1732">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1732">String</span></span>||<span data-ttu-id="c3a7b-1733">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1733">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="c3a7b-1734">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1734">Object</span></span>|<span data-ttu-id="c3a7b-1735">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1735">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1736">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1736">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a7b-1737">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1737">Object</span></span>|<span data-ttu-id="c3a7b-1738">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1738">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1739">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1739">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c3a7b-1740">function</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1740">function</span></span>|<span data-ttu-id="c3a7b-1741">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1741">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1742">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1742">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c3a7b-1743">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1743">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c3a7b-1744">エラー</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1744">Errors</span></span>

|<span data-ttu-id="c3a7b-1745">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1745">Error code</span></span>|<span data-ttu-id="c3a7b-1746">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1746">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="c3a7b-1747">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1747">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1748">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1748">Requirements</span></span>

|<span data-ttu-id="c3a7b-1749">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1749">Requirement</span></span>|<span data-ttu-id="c3a7b-1750">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1750">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1751">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1751">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1752">1.1</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1752">1.1</span></span>|
|[<span data-ttu-id="c3a7b-1753">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1753">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1754">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1754">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3a7b-1755">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1755">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1756">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1756">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-1757">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1757">Example</span></span>

<span data-ttu-id="c3a7b-1758">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1758">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="c3a7b-1759">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1759">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="c3a7b-1760">サポートされているイベントの種類のイベント ハンドラーを削除します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1760">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="c3a7b-1761">現在、サポートされて`Office.EventType.AttachmentsChanged`いる`Office.EventType.AppointmentTimeChanged`イベント`Office.EventType.EnhancedLocationsChanged`の`Office.EventType.RecipientsChanged`種類は`Office.EventType.RecurrenceChanged`、、、、、です。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1761">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a7b-1762">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1762">Parameters</span></span>

| <span data-ttu-id="c3a7b-1763">名前</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1763">Name</span></span> | <span data-ttu-id="c3a7b-1764">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1764">Type</span></span> | <span data-ttu-id="c3a7b-1765">属性</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1765">Attributes</span></span> | <span data-ttu-id="c3a7b-1766">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1766">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c3a7b-1767">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1767">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c3a7b-1768">ハンドラーを取り消すイベント。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1768">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="c3a7b-1769">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1769">Object</span></span> | <span data-ttu-id="c3a7b-1770">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1770">&lt;optional&gt;</span></span> | <span data-ttu-id="c3a7b-1771">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1771">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c3a7b-1772">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1772">Object</span></span> | <span data-ttu-id="c3a7b-1773">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1773">&lt;optional&gt;</span></span> | <span data-ttu-id="c3a7b-1774">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1774">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c3a7b-1775">function</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1775">function</span></span>| <span data-ttu-id="c3a7b-1776">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1776">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1777">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1777">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1778">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1778">Requirements</span></span>

|<span data-ttu-id="c3a7b-1779">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1779">Requirement</span></span>| <span data-ttu-id="c3a7b-1780">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1780">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1781">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1781">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3a7b-1782">1.7</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1782">1.7</span></span> |
|[<span data-ttu-id="c3a7b-1783">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1783">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3a7b-1784">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1784">ReadItem</span></span> |
|[<span data-ttu-id="c3a7b-1785">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1785">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c3a7b-1786">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1786">Compose or Read</span></span> |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="c3a7b-1787">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1787">saveAsync([options], callback)</span></span>

<span data-ttu-id="c3a7b-1788">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1788">Asynchronously saves an item.</span></span>

<span data-ttu-id="c3a7b-1789">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1789">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="c3a7b-1790">Outlook on the web またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1790">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="c3a7b-1791">キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1791">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a7b-1792">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1792">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="c3a7b-1793">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1793">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="c3a7b-p198">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p198">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="c3a7b-1797">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1797">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="c3a7b-1798">Outlook on Mac では、会議の保存はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1798">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="c3a7b-1799">`saveAsync` メソッドは、作成モードの会議から呼び出されると失敗します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1799">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="c3a7b-1800">回避策については、「[Office JS API を使用して Outlook for Mac で会議を下書きとして保存できない](https://support.microsoft.com/help/4505745)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1800">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="c3a7b-1801">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1801">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a7b-1802">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1802">Parameters</span></span>

|<span data-ttu-id="c3a7b-1803">名前</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1803">Name</span></span>|<span data-ttu-id="c3a7b-1804">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1804">Type</span></span>|<span data-ttu-id="c3a7b-1805">属性</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1805">Attributes</span></span>|<span data-ttu-id="c3a7b-1806">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1806">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c3a7b-1807">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1807">Object</span></span>|<span data-ttu-id="c3a7b-1808">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1808">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1809">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1809">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a7b-1810">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1810">Object</span></span>|<span data-ttu-id="c3a7b-1811">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1811">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1812">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1812">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c3a7b-1813">function</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1813">function</span></span>||<span data-ttu-id="c3a7b-1814">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1814">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c3a7b-1815">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1815">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1816">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1816">Requirements</span></span>

|<span data-ttu-id="c3a7b-1817">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1817">Requirement</span></span>|<span data-ttu-id="c3a7b-1818">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1818">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1819">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1819">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1820">1.3</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1820">1.3</span></span>|
|[<span data-ttu-id="c3a7b-1821">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1821">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1822">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1822">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3a7b-1823">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1823">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1824">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1824">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c3a7b-1825">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1825">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="c3a7b-p200">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p200">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="c3a7b-1828">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1828">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="c3a7b-1829">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1829">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="c3a7b-p201">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p201">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c3a7b-1833">パラメーター</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1833">Parameters</span></span>

|<span data-ttu-id="c3a7b-1834">名前</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1834">Name</span></span>|<span data-ttu-id="c3a7b-1835">型</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1835">Type</span></span>|<span data-ttu-id="c3a7b-1836">属性</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1836">Attributes</span></span>|<span data-ttu-id="c3a7b-1837">説明</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1837">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="c3a7b-1838">String</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1838">String</span></span>||<span data-ttu-id="c3a7b-p202">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-p202">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="c3a7b-1842">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1842">Object</span></span>|<span data-ttu-id="c3a7b-1843">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1843">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1844">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1844">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c3a7b-1845">Object</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1845">Object</span></span>|<span data-ttu-id="c3a7b-1846">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1846">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1847">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1847">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="c3a7b-1848">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1848">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="c3a7b-1849">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1849">&lt;optional&gt;</span></span>|<span data-ttu-id="c3a7b-1850">`text` の場合、Outlook on the web とデスクトップ クライアントでは現在のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1850">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="c3a7b-1851">フィールドが HTML エディターの場合、データが HTML の場合でもテキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1851">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="c3a7b-1852">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Outlook on the web では現在のスタイルが適用され、Outlook デスクトップ クライアントでは既定のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1852">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="c3a7b-1853">フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1853">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="c3a7b-1854">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1854">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="c3a7b-1855">function</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1855">function</span></span>||<span data-ttu-id="c3a7b-1856">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1856">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3a7b-1857">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1857">Requirements</span></span>

|<span data-ttu-id="c3a7b-1858">要件</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1858">Requirement</span></span>|<span data-ttu-id="c3a7b-1859">値</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1859">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3a7b-1860">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1860">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c3a7b-1861">1.2</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1861">1.2</span></span>|
|[<span data-ttu-id="c3a7b-1862">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1862">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c3a7b-1863">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1863">ReadWriteItem</span></span>|
|[<span data-ttu-id="c3a7b-1864">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1864">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="c3a7b-1865">作成</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1865">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c3a7b-1866">例</span><span class="sxs-lookup"><span data-stu-id="c3a7b-1866">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
